<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<%
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
%>
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

' 권한관리 추가 
Dim lgBizAreaAuthYn, lgAuthBizAreaCd, lgAuthBizAreaNm		' 사업장 
Dim lgInternalAuthYn, lgInternalCd					' 내부부서 
Dim lgSubInternalAuthYn, lgSubInternalCd			' 내부부서(하위포함)
Dim lgAuthUsrIDAuthYn, lgAuthUsrID					' 개인 
Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalAuthCd, lgAuthUsrIDAuthSQL


'-- eWare If Begin
    If Trim(gEware) <> "" Then
		Dim sAppFg
		Dim strTemp
		Dim strIf

		lgErrorStatus     = "NO"
		lgErrorPos        = ""                                                       '☜: Set to space
    End If
'-- eWare If End

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
             Call SubBizDeleteMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear    
	
	Const C_ATempGl_gl_no = 0
	Const C_ATempGlItem_item_seq = 1
	
	Const A352_E1_a_temp_gl_temp_gl_no = 0
	Const A352_E1_a_temp_gl_temp_gl_dt = 1
	Const A352_E1_a_temp_gl_gl_type = 2
	Const A352_E1_a_temp_gl_input_type = 3
	Const A352_E1_a_temp_gl_cr_amt = 4
	Const A352_E1_a_temp_gl_cr_loc_amt = 5
	Const A352_E1_a_temp_gl_dr_amt = 6
	Const A352_E1_a_temp_gl_dr_loc_amt = 7
	Const A352_E1_a_temp_gl_conf_fg = 8
	Const A352_E1_a_temp_gl_temp_gl_desc = 9
	Const A352_E1_a_temp_gl_project_no = 10
	Const A352_E1_a_temp_gl_org_change_id = 11
	Const A352_E1_a_temp_gl_dept_cd = 12
	Const A352_E1_dept_nm = 13
	
	Const A352_EG1_a_temp_gl_item_item_seq		= 0
	Const A352_EG1_a_temp_gl_item_dept_cd		= 1
	Const A352_EG1_a_temp_gl_item_dept_nm		= 2
	Const A352_EG1_a_temp_gl_item_acct_cd		= 3
	Const A352_EG1_a_temp_gl_item_acct_nm		= 4
	Const A352_EG1_a_temp_gl_item_dr_cr_fg		= 5	
	Const A352_EG1_a_temp_gl_item_item_amt		= 6
	Const A352_EG1_a_temp_gl_item_item_loc_amt	= 7
	Const A352_EG1_a_temp_gl_item_vat_type		= 8
	Const A352_EG1_a_temp_gl_item_item_desc		= 9
	Const A352_EG1_a_temp_gl_item_xch_rate		= 10
	Const A352_EG1_a_temp_gl_item_doc_cur		= 11
	Const A352_EG1_a_temp_gl_item_project_no	= 12
	Const A352_EG1_a_temp_gl_item_gl_no			= 13
	Const A352_EG1_a_temp_gl_item_org_change_id	= 14
	Const A352_EG1_a_temp_gl_item_acct_type		= 15

	Dim PAGG005_cALkUpTmpGlSvr
	Dim iStrData
    Dim iexportData
    Dim iexportData1
    Dim iLngRow
    Dim iLngCol
    Dim iStrCurrency
        
    Dim iArrKey
    ReDim iArrKey(5)
    
	On Error Resume Next
    Err.Clear    

    iArrKey(C_ATempGl_gl_no) = UCase(Trim(Request("txtTempGlNo")))
    iArrKey(C_ATempGlItem_item_seq) = "0"
    
    iArrKey(C_ATempGlItem_item_seq+1) = lgAuthBizAreaCd
    iArrKey(C_ATempGlItem_item_seq+2) = lgInternalCd
    iArrKey(C_ATempGlItem_item_seq+3) = lgSubInternalCd
    iArrKey(C_ATempGlItem_item_seq+4) = lgAuthUsrID

    Set PAGG005_cALkUpTmpGlSvr = Server.CreateObject("PAGG005.cALkUpTmpGlSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write " With Parent " & vbCr				
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"")"	       & vbCr
		Response.Write " End With " & vbCr
		Response.Write "</Script> " & vbCr    
		Exit Sub
    End If
    
	Call PAGG005_cALkUpTmpGlSvr.A_LOOKUP_TEMP_GL_SVR(gStrGlobalCollection, iArrKey, iexportData, iexportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write " With Parent " & vbCr				
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"")"	       & vbCr
		Response.Write " End With " & vbCr
		Response.Write "</Script> " & vbCr	
		Set PAGG005_cALkUpTmpGlSvr = Nothing        
		Exit Sub
    End If    

    Set PAGG005_cALkUpTmpGlSvr = Nothing
    
    iStrData = ""	
	For iLngRow = 0 To UBound(iexportData1, 1) 
		iStrCurrency = ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_doc_cur))

		iStrData = iStrData & Chr(11) & iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_seq)
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dept_cd)))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dept_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_cd))
		iStrData = iStrData & Chr(11) & ""		
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_nm))
		iStrData = iStrData & Chr(11) & UCase(ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dr_cr_fg)))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_doc_cur))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & UNINumClientFormat(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_xch_rate),	ggExchRate.DecPoint,	 0)' UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_xch_rate),iStrCurrency,ggExchRateNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_loc_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_desc))
		iStrData = iStrData & Chr(11) & UCase(Trim(ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_vat_type))))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_cd))
		iStrData = iStrData & Chr(11) & iLngRow + 1
		
		iStrData = iStrData & Chr(11) & Chr(12)		
	Next

'-- eWare If Begin
    If Trim(gEware) <> "" Then		
		strTemp = UCase(Trim(iexportData(A352_E1_a_temp_gl_temp_gl_no))) & gColSep	
		sAppFg = SubQuery_If("eWare",strTemp )				
    End If    
'-- eWare If End


	'++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'기표부서 검색 함수 추가 >>air
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Call fnLoginDeptCd(UCase(Trim(Request("txtTempGlNo"))))
		
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent " & vbCr		
	Response.Write " 	.frm1.txtTempGlNo.value			= """ & ConvSPChars(UCase(Trim(iexportData(A352_E1_a_temp_gl_temp_gl_no))))			& """" & vbCr
	Response.Write "	.frm1.txtTempGLDt.Text			= """ & UNIDateClientFormat(iexportData(A352_E1_a_temp_gl_temp_gl_dt))				& """" & vbCr
	Response.Write " 	.frm1.cboGLtype.value			= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_gl_type))							& """" & vbCr
	'Response.Write " 	.frm1.txtGlinputType.value		= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_input_type))						& """" & vbCr
	Response.Write " 	.frm1.cboGlInputType.Value		= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_input_type))						& """" & vbCr
'	Response.Write " 	.frm1.txtCrAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_cr_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr
	Response.Write " 	.frm1.txtCrLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
'	Response.Write " 	.frm1.txtDrAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_dr_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr 
	Response.Write " 	.frm1.txtDrLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr							
	Response.Write " 	.frm1.txtDesc.Value				= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_temp_gl_desc))					& """" & vbCr
	Response.Write " 	.frm1.txtDeptCd.value			= """ & ConvSPChars(UCase(Trim(iexportData(A352_E1_a_temp_gl_dept_cd))))			& """" & vbCr
	Response.Write " 	.frm1.txtDeptNm.value			= """ & ConvSPChars(iexportData(A352_E1_dept_nm))									& """" & vbCr
	Response.Write " 	.frm1.cboConfFg.value			= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_conf_fg))							& """" & vbCr
	Response.Write " 	.frm1.hCongFg.value				= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_conf_fg))							& """" & vbCr	
'-- eWare If Begin
    If Trim(gEware) <> "" Then
		Response.Write ".frm1.cboAppFg.value		    = """ & ConvSPChars(UCase(Trim(sAppFg)))											& """" & vbCr
    End If    		
'-- eWare If End	
'	Response.Write " 	.frm1.txtDocCur.value			= """ & ConvSPChars(iStrCurrency)													& """" & vbCr
	Response.Write " 	.frm1.htxtTempGlNo.value		= """ & ConvSPChars(UCase(Trim(iexportData(A352_E1_a_temp_gl_temp_gl_no))))			& """" & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	      " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData   & """ ,""F""" & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow  & ",.C_DocCur,.C_ItemAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow  & ",.C_DocCur,.C_Exchrate,  ""D"" ,""I"",""X"",""X"")" & vbCr	
	Response.Write " 	.DbQueryOk								  " & vbCr
	Response.Write " End With " & vbCr
	Response.Write "</Script> " & vbCr 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next	
	Err.Clear
	
	Const A393_I2_a_temp_gl_temp_gl_no = 0
    Const A393_I2_a_temp_gl_temp_gl_dt = 1
    Const A393_I2_a_temp_gl_org_change_id = 2
    Const A393_I2_a_temp_gl_dept_cd = 3
    Const A393_I2_a_temp_gl_gl_type = 4
    Const A393_I2_a_temp_gl_gl_input_type = 5
    Const A393_I2_a_temp_gl_temp_gl_desc = 6
    Const A393_I2_a_temp_gl_project_no = 7

	Dim PAGG005_cAMngTmpGlSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetTempGlNo

	Dim I4_a_data_auth ' 권한관리추가 2006-08-01 jyk

	Const A393_I4_a_data_auth_data_BizAreaCd = 0
	Const A393_I4_a_data_auth_data_internal_cd = 1
	Const A393_I4_a_data_auth_data_sub_internal_cd = 2
	Const A393_I4_a_data_auth_data_auth_usr_id = 3
	
	Dim iLngMaxRow
	Dim iLngMaxRow3
	Dim iLngRow
	Dim iArrTemp1
	Dim iArrTemp2
	Dim retVal
	
	iLngMaxRow  = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    iLngMaxRow3 = CInt(Request("txtMaxRows3"))
	
	'--------------------------------------------------------------------
	'A_GL에 대한 정보  Setting
	'--------------------------------------------------------------------
	iCommandSent = Request("txtCommandMode")	'Spread Sheet 내용을 담고 있는 Element명 
	I1_b_currency = gCurrency

    ReDim I2_a_gl(7)
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)		= UCase(Trim(Request("txtTempGlNo")))
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_dt)		= UNIConvDate(Request("txtTempGlDt"))
	I2_a_gl(A393_I2_a_temp_gl_org_change_id)	= Request("hOrgChangeId")
	I2_a_gl(A393_I2_a_temp_gl_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A393_I2_a_temp_gl_gl_type)			= Request("cboGlType") 
	I2_a_gl(A393_I2_a_temp_gl_gl_input_type)	= Request("cboGlInputType")      
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_desc)		= Request("txtDesc")

       I2_a_gl(A393_I2_a_temp_gl_project_no)	        = Request("txthpjt_no")

	Redim I4_a_data_auth(3)
	I4_a_data_auth(A393_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	'--------------------------------------------------------------------
	'A_GL_ITEM에 대한 정보  Setting
	'--------------------------------------------------------------------
    iArrTemp1 = Split(Request("txtSpread"), gRowSep)	'ITEM SPREAD	

   	For iLngRow = 1 To iLngMaxRow
		iArrTemp2 = Split(iArrTemp1(iLngRow-1), gColSep)
        txtSpread = txtSpread & "C" & gColSep
		' 전체 삭제후 생성될 내용(Update, Insert) 만 전달 
		If iArrTemp2(0) <> "D" Then		
		    txtSpread = txtSpread & Cint(iArrTemp2(1))										& gColSep		'current row
		    txtSpread = txtSpread & Cint(iArrTemp2(2))										& gColSep		'ItemSEQ  * Key
			txtSpread = txtSpread & iArrTemp2(4)											& gColSep		'계정코드 
			txtSpread = txtSpread & iArrTemp2(5)											& gColSep		'차대구분 
			txtSpread = txtSpread & Request("hOrgChangeId")	& gColSep			
			txtSpread = txtSpread & iArrTemp2(3)			& gColSep		'부서	
			txtSpread = txtSpread & iArrTemp2(11)			& gColSep   '거래통화			
			
			If Trim(iArrTemp2(9)) = "" then
				txtSpread = txtSpread & "0"													& gColSep		'환율 
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(9))									& gColSep		
			End if
			
			txtSpread = txtSpread & iArrTemp2(10)											& gColSep		'부가세 type			
     	    
     	    If Trim(iArrTemp2(6)) = "" then																	'거래금액 
				txtSpread = txtSpread & ""													& gColSep	
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(6))									& gColSep	
			End if			
			If Trim(iArrTemp2(7)) = "" then								
				txtSpread = txtSpread & ""													& gColSep					
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(7))									& gColSep		
			End if
			
			' 200909.by lws add
			
			txtSpread = txtSpread & iArrTemp2(8)	& gColSep		'비고 12
			txtSpread = txtSpread & ""	& gColSep		' 13
			txtSpread = txtSpread & ""	& gColSep		' 14
			txtSpread = txtSpread & ""	& gColSep		' 15
			txtSpread = txtSpread & "0"	& gColSep		' 16
			txtSpread = txtSpread & "0"	& gColSep		' 17
			txtSpread = txtSpread &  iLngRow 	& gRowSep		' 18
				
		End If
	Next


    '--------------------------------------------------------------------
	'A_GL_DTL에 대한 정보  Setting
	'--------------------------------------------------------------------
	iArrTemp1 = Split(Request("txtSpread3"), gRowSep)

	For iLngRow = 1 to iLngMaxRow3
		iArrTemp2 = Split(iArrTemp1(iLngRow-1), gColSep)
		txtSpread3 = txtSpread3 & "C" & gColSep
		If iArrTemp2(0) <> "D" Then
		        txtSpread3 = txtSpread3 & Cint(iArrTemp2(1))	& gColSep
		        txtSpread3 = txtSpread3 & Cint(iArrTemp2(2))	& gColSep
		        txtSpread3 = txtSpread3 & Trim(iArrTemp2(3))	& gColSep
		    	txtSpread3 = txtSpread3 & UCase(iArrTemp2(4))	& gRowSep		    
		End If
   	Next

	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------	
	Set PAGG005_cAMngTmpGlSvr = Server.CreateObject("PAGG005.cAMngTmpGlSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iStrRetTempGlNo = PAGG005_cAMngTmpGlSvr.A_MANAGE_TEMP_GL_SVR(gStrGlobalCollection, iCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3,,I4_a_data_auth) 	

	If CheckSYSTEMError(Err, True) = True Then		
		Set PAGG005_cAMngTmpGlSvr = Nothing
		Exit Sub
    End If

    Set PAGG005_cAMngTmpGlSvr  = Nothing

    If InStr(1,ConvSPChars(UCase(Trim(Request("txtTempGlNo")))),"'") > 0 Then
		iStrRetTempGlNo = ConvSPChars(UCase(Trim(Request("txtTempGlNo"))))
	End If

'-- eWare If Begin
	If Trim(gEware) <> "" Then
		strIf = ""
		strIf = iStrRetTempGlNo &  gColSep
		'strIf = strIf & Request("CboAppFg") &  gColSep
		strIf = strIf & "U" &  gColSep
 		Call SubDelete_If("eWare",strIf)
		Call SubSave_if("eWare",strIf)
	End If
'-- eWare If End

	Response.Write " <Script Language=vbscript>											" & vbCr
	Response.Write " With parent														" & vbCr
    Response.Write "	.DbSaveOk """ & ConvSPChars(iStrRetTempGlNo)	&			 """" & vbCr    
    Response.Write " End With															" & vbCr
    Response.Write " </Script>															" & vbCr
End Sub    

'============================================================================================================
' Name : SubBizDeleteMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDeleteMulti()
	On Error Resume Next
	Err.Clear

	Const A393_I2_a_temp_gl_temp_gl_no = 0
    Const A393_I2_a_temp_gl_temp_gl_dt = 1
    Const A393_I2_a_temp_gl_org_change_id = 2
    Const A393_I2_a_temp_gl_dept_cd = 3
    Const A393_I2_a_temp_gl_gl_type = 4
    Const A393_I2_a_temp_gl_gl_input_type = 5
    Const A393_I2_a_temp_gl_temp_gl_desc = 6
    Const A393_I2_a_temp_gl_project_no = 7

	Dim PAGG005_cAMngTmpGlSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetTempGlNo

	Dim I4_a_data_auth ' 권한관리추가 2006-08-01 jyk

	Const A393_I4_a_data_auth_data_BizAreaCd = 0
	Const A393_I4_a_data_auth_data_internal_cd = 1
	Const A393_I4_a_data_auth_data_sub_internal_cd = 2
	Const A393_I4_a_data_auth_data_auth_usr_id = 3

	iCommandSent = "DELETE"
	ReDim I2_a_gl(7)
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)			= UCase(Trim(Request("txtTempGlNo")))
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_dt)			= UNIConvDate(Request("txtTempGlDt"))
	I2_a_gl(A393_I2_a_temp_gl_org_change_id)		= Trim(Request("txtOrgChangeId"))
	I2_a_gl(A393_I2_a_temp_gl_dept_cd)				= UCase(Trim(Request("txtDeptCd")))
   'I2_a_gl(A393_I2_a_temp_gl_gl_type)				= Trim(Request("cboGlType")) 
	I2_a_gl(A393_I2_a_temp_gl_gl_input_type)		= Trim(Request("txtGlinputType"))      
   'I2_a_gl(A393_I2_a_temp_gl_temp_gl_desc)			= Request("txtDesc") 

	Redim I4_a_data_auth(3)
	I4_a_data_auth(A393_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A393_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	Set PAGG005_cAMngTmpGlSvr = Server.CreateObject("PAGG005.cAMngTmpGlSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iStrRetTempGlNo = PAGG005_cAMngTmpGlSvr.A_MANAGE_TEMP_GL_SVR(gStrGlobalCollection, iCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3,,I4_a_data_auth)

	If CheckSYSTEMError(Err, True) = True Then
		Set PAGG005_cAMngTmpGlSvr = Nothing
		Exit Sub
    End If

    Set PAGG005_cAMngTmpGlSvr  = Nothing

'-- eWare If Begin
    If Trim(gEware) <> "" Then
		strIf = ""
		strIf = UCase(Trim(Request("txtTempGlNo"))) &  gColSep
		Call SubDelete_if("eWare",strIf)
    End If	
'-- eWare If End

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.DbDeleteOK										" & vbCr    
    Response.Write " End With											" & vbCr
    Response.Write " </Script>											" & vbCr
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()

End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()

End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()

End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                               '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "SC"
            If CheckSYSTEMError(pErr,True) = True Then
				ObjectContext.SetAbort
				Call SetErrorStatus
            Else
				If CheckSQLError(pConn,True) = True Then
					ObjectContext.SetAbort
					Call SetErrorStatus
				End If
            End If
        Case "SD"
            If CheckSYSTEMError(pErr,True) = True Then
				ObjectContext.SetAbort
				Call SetErrorStatus
            Else
				If CheckSQLError(pConn,True) = True Then
					ObjectContext.SetAbort
					Call SetErrorStatus
				End If
            End If
        Case "SR"
            If CheckSYSTEMError(pErr,True) = True Then
				ObjectContext.SetAbort
				Call SetErrorStatus
            Else
				If CheckSQLError(pConn,True) = True Then
				   ObjectContext.SetAbort
				   Call SetErrorStatus
				End If
            End If
        Case "SU"
			If CheckSYSTEMError(pErr,True) = True Then
			    ObjectContext.SetAbort
			    Call SetErrorStatus
			Else
			    If CheckSQLError(pConn,True) = True Then
					ObjectContext.SetAbort
					Call SetErrorStatus
			    End If
			End If
    End Select
End Sub

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
'-- eWare if Begin
Function SubQuery_If(sCase,strQuery)
	Dim  sReturnVal
	Dim  arrVal_e
	DIm  ReturnQ			

	arrVal_e = Split(strQuery, gColSep)

	Select Case sCase
		Case "eWare"
			If  CommonQueryRs(" APP_FG ","  X_A_TEMP_GL_IF "," TEMP_GL_NO = " & FilterVar(arrVal_e(0), "''", "S"),ReturnQ,"","","","","","") Then
				If ReturnQ = chr(11) Or isnull(ReturnQ) Or ReturnQ = "X" Then
					sReturnVal = "U"
				Else
					arrVal_e = Split(ReturnQ, Chr(11))
					sReturnVal = arrVal_e(0)
				End If
			Else
				sReturnVal = "U"
			End If
	End Select

	SubQuery_If = 	sReturnVal
End Function

'============================================================================================================
' Name : SubSave_If
' Desc : 
'============================================================================================================
Sub SubSave_If(sCase,strSave)
	Dim	arrVal_q

	arrVal_q = Split(strSave, gColSep)

	Call SubOpenDB(lgObjConn)																'☜: Make a DB Connection

	lgStrSQL =  ""
	Select Case sCase
		Case "eWare"
			lgStrSQL = "INSERT INTO X_A_TEMP_GL_IF "
			lgStrSQL = lgStrSQL & " ( TEMP_GL_NO	, "										'1 전표번호 
			lgStrSQL = lgStrSQL & "   APP_FG	, "											'2 E-WARE승인 여부 
			lgStrSQL = lgStrSQL & "   INSRT_DT	, "											'0 등록일 
			lgStrSQL = lgStrSQL & "   INSRT_USER_ID	, "										'1 등록자 
			lgStrSQL = lgStrSQL & "   UPDT_DT	, "											'2 수정일 
			lgStrSQL = lgStrSQL & "   UPDT_USER_ID	) "										'3 수정자 
			lgStrSQL = lgStrSQL & " VALUES(" & FilterVar(arrVal_q(0), "''", "S")  & ", "		'1 전표번호 
			lgStrSQL = lgStrSQL & FilterVar(arrVal_q(1)	, "''", "S")			& ", "		'2 E-WARE승인 여부 
			lgStrSQL = lgStrSQL &	 "getdate(), " 											'0
			lgStrSQL = lgStrSQL &	FilterVar(gUsrID	, "''", "S")			    & ", "		'1
			lgStrSQL = lgStrSQL &	 "getdate(), " 											'2
			lgStrSQL = lgStrSQL &	FilterVar(gUsrID	, "''", "S")			    & ")" 		'3
	End Select

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords								'상환테이블인서트 
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)

	Call SubCloseDB(lgObjConn)																'☜: Close DB Connection
End Sub

'============================================================================================================
' Name : SubDelete_If
' Desc : 
'============================================================================================================
Sub SubDelete_If(sCase,strSave)
	Dim	arrVal_q

	arrVal_q = Split(strSave, gColSep)

	Call SubOpenDB(lgObjConn)																'☜: Make a DB Connection

	lgStrSQL = ""

	Select Case sCase
		Case "eWare"
			lgStrSQL = "DELETE FROM X_A_TEMP_GL_IF "
			lgStrSQL = lgStrSQL & " WHERE TEMP_GL_NO = " & FilterVar(arrVal_q(0)	, "''", "S")
	End Select

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords								'상환테이블인서트 
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)

	Call SubCloseDB(lgObjConn)																'☜: Close DB Connection
End Sub
'-- eWare If End




Sub fnLoginDeptCd(strTempGlNo)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

		Call SubOpenDB(lgObjConn)	

		lgStrSQL = ""
		lgStrSQL = lgStrSQL & " SELECT ORG_NM "
		lgStrSQL = lgStrSQL & " FROM   Z_USR_ORG_MAST(NOLOCK) "
		lgStrSQL = lgStrSQL & " WHERE  ORG_TYPE = 'DP' "
		lgStrSQL = lgStrSQL & "        AND USE_YN = 'Y' "
		lgStrSQL = lgStrSQL & "        AND USR_ID = (SELECT	TOP 1 INSRT_USER_ID "
		lgStrSQL = lgStrSQL & " 					FROM	A_TEMP_GL(NOLOCK) "
		lgStrSQL = lgStrSQL & " 					WHERE	TEMP_GL_NO = " & FilterVar(strTempGlNo,"''", "S") & ")"

 'Call ServerMesgBox("a" , vbInformation, I_MKSCRIPT) 
       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
          Response.Write  " <Script Language=vbscript>            " & vbCr
          Response.Write  "   Parent.Frm1.txtLoginDeptNm.Value  = """ & lgObjRs(0) & """" & vbCr             
          Response.Write  " </Script> " & vbCr
          Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
       'Else
   	
       '   Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
       End If
End Sub		

%>

