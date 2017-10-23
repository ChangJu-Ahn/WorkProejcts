<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5407mb1
'*  4. Program Name         : 미결반제(신용카드)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2002/11/05
'*  8. Modified date(Last)  : 2003/08/12
'*  9. Modifier (First)     : KIM HO YOUNG
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>

<%
'#########################################################################################################
'												2. 조건부 
'##########################################################################################################

													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd	

On Error Resume Next														'☜: 
Err.Clear

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim lgOpModeCRUD
Dim I3_trans_type

	I3_trans_type = "AP011"

    lgOpModeCRUD = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDeleteMulti()
    End Select

'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear

	Const A351_I1_a_cls_no = 0
	
	Const A351_E1_a_temp_gl_temp_gl_no = 0
	Const A351_E1_a_temp_gl_temp_gl_dt = 1
	Const A351_E1_a_temp_gl_gl_type = 2
	Const A351_E1_a_temp_gl_input_type = 3
	Const A351_E1_a_temp_gl_cr_amt = 4
	Const A351_E1_a_temp_gl_cr_loc_amt = 5
	Const A351_E1_a_temp_gl_dr_amt = 6
	Const A351_E1_a_temp_gl_dr_loc_amt = 7
	Const A351_E1_a_temp_gl_temp_gl_desc = 8
	Const A351_E1_a_temp_gl_project_no = 9
	Const A351_E1_a_temp_gl_org_change_id = 10
	Const A351_E1_a_temp_gl_dept_cd = 11
	Const A351_E1_dept_nm = 12
	Const A351_E1_a_temp_gl_hq_brch_fg = 13
    Const A351_E1_a_temp_gl_hq_brch_no = 14	
    Const A351_E1_a_temp_gl_conf_fg = 15
    
	Const A351_E2_a_conf_fg = 0
	Const A351_E2_a_Temp_gl_no = 1  
	Const A351_E2_a_gl_no = 2
	
    
	Const A351_EG21_a_mgnt_val1 = 0
	Const A351_EG21_a_mgnt_val2 = 1
	Const A351_EG21_a_user = 2
	Const A351_EG21_gl_no = 3
	Const A351_EG21_a_gl_dt = 4
	Const A351_EG21_a_doc_cur = 5
	Const A351_EG21_a_xch_rate = 6
	Const A351_EG21_a_open_doc_amt = 7
	Const A351_EG21_a_open_amt = 8
	Const A351_EG21_a_dept_cd = 9
	Const A351_EG21_a_dept_nm = 10
	Const A351_EG21_a_acct_cd = 11
	Const A351_EG21_a_acct_nm = 12
	Const A351_EG21_a_dr_cr_fg = 13
	Const A351_EG21_a_dr_cr_nm = 14
	Const A351_EG21_a_cardco_nm = 15
	Const A351_EG21_a_opengl_item_seq = 16
	
	Const A351_EG11_a_temp_gl_item_item_seq		= 0
	Const A351_EG11_a_temp_gl_item_dept_cd		= 1
	Const A351_EG11_a_temp_gl_item_dept_nm		= 2
	Const A351_EG11_a_temp_gl_item_acct_cd		= 3
	Const A351_EG11_a_temp_gl_item_acct_nm		= 4
	Const A351_EG11_a_temp_gl_item_dr_cr_fg		= 5	
	Const A351_EG11_a_temp_gl_item_item_amt		= 6
	Const A351_EG11_a_temp_gl_item_item_loc_amt	= 7
	Const A351_EG11_a_temp_gl_item_vat_type		= 8
	Const A351_EG11_a_temp_gl_item_item_desc		= 9
	Const A351_EG11_a_temp_gl_item_xch_rate		= 10
	Const A351_EG11_a_temp_gl_item_doc_cur		= 11
	Const A351_EG11_a_temp_gl_item_project_no		= 12
	Const A351_EG11_a_temp_gl_item_gl_no			= 13
	Const A351_EG11_a_temp_gl_item_org_change_id	= 14
	Const A351_EG11_a_temp_gl_item_acct_type		= 15
    
	Dim PAUG035_cALkUpClsCardSvr
	Dim iStrData
	Dim iStrData1
	Dim iStrData2
    Dim iexportData
    Dim iexportData1
    Dim iexportData2
    Dim E2_conf_fg
    Dim iLngRow
    Dim iLngCol
    Dim iStrCurrency
    Dim iStrCurrency1

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
    
    Dim iStrClsNo

    Redim iStrClsNo(A351_I1_a_cls_no+4)
    
	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))    
    
    iStrClsNo(A351_I1_a_cls_no)   = UCase(Trim(Request("txtClsNo")))
	iStrClsNo(A351_I1_a_cls_no+1) = lgAuthBizAreaCd
	iStrClsNo(A351_I1_a_cls_no+2) = lgInternalCd
	iStrClsNo(A351_I1_a_cls_no+3) = lgSubInternalCd
	iStrClsNo(A351_I1_a_cls_no+4) = lgAuthUsrID    
    
    Set PAUG035_cALkUpClsCardSvr = Server.CreateObject("PAUG035.cALkUpClsCardSvr")
    
    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
    
	Call PAUG035_cALkUpClsCardSvr.A_LOOKUP_CLS_CARD_SVR(gStrGlobalCollection, iStrClsNo, iexportData, E2_conf_fg,  iexportData1, iexportData2)

	If CheckSYSTEMError(Err, True) = True Then					
		Set PAUG035_cALkUpClsCardSvr = Nothing        
		Exit Sub
    End If    

    Set PAUG035_cALkUpClsCardSvr = Nothing
    
	iStrCurrency = ConvSPChars(iexportData1(0, A351_EG21_a_doc_cur))
    iStrData	= ""	
    iStrData1	= ""
    iStrData2	= ""

	For iLngRow = 0 To UBound(iexportData2, 1) 
		iStrData = iStrData & Chr(11) & iexportData2(iLngRow, A351_EG21_a_opengl_item_seq)
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_mgnt_val1))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_mgnt_val2))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_user))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_cardco_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_gl_no))
		iStrData = iStrData & Chr(11) & UniDateClientFormat(iexportData2(iLngRow, A351_EG21_a_gl_dt))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_doc_cur))
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData2(iLngRow, A351_EG21_a_xch_rate),gCurrency,ggAmtOfMoneyNo, "X" , "X")		
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData2(iLngRow, A351_EG21_a_open_doc_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData2(iLngRow, A351_EG21_a_open_amt),gCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_dept_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_dept_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_acct_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_acct_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_dr_cr_fg))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData2(iLngRow, A351_EG21_a_opengl_item_seq))
		iStrData = iStrData & Chr(11) & iLngRow + 1
		iStrData = iStrData & Chr(11) & Chr(12)		
	Next

	For iLngRow = 0 To UBound(iexportData1, 1) 
		iStrCurrency1 = ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_doc_cur))

		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_item_seq))
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_dept_cd))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_dept_nm))
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_acct_cd))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_acct_nm))
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_dr_cr_fg))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_doc_cur))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_xch_rate),gCurrency,ggAmtOfMoneyNo, "X" , "X")		
		iStrData1 = iStrData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData1 = iStrData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_item_loc_amt),gCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_item_desc))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_gl_no))
		iStrData1 = iStrData1 & Chr(11) & ConvSPChars(iexportData1(iLngRow, A351_EG11_a_temp_gl_item_item_seq))
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & ""
		iStrData1 = iStrData1 & Chr(11) & iLngRow + 1
		iStrData1 = iStrData1 & Chr(11) & Chr(12)	
	Next
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent " & vbCr		
	Response.Write "	.frm1.txtClsDt.Text			= """ & UNIDateClientFormat(iexportData(A351_E1_a_temp_gl_temp_gl_dt))	& """" & vbCr
	Response.Write " 	.frm1.cboGLtype.value		= """ & iexportData(A351_E1_a_temp_gl_gl_type)							& """" & vbCr
	Response.Write " 	.frm1.cboGlInputType.Value	= """ & ConvSPChars(iexportData(A351_E1_a_temp_gl_input_type))			& """" & vbCr
	Response.Write " 	.frm1.txtCrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A351_E1_a_temp_gl_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write " 	.frm1.txtCrLocAmt2.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A351_E1_a_temp_gl_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write " 	.frm1.txtDrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A351_E1_a_temp_gl_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr							
	Response.Write " 	.frm1.txtDrLocAmt2.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A351_E1_a_temp_gl_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr								
	Response.Write " 	.frm1.txtDeptCd.value		= """ & UCase(Trim(iexportData(A351_E1_a_temp_gl_dept_cd)))				& """" & vbCr
	Response.Write " 	.frm1.txtDeptNm.value		= """ & ConvSPChars(iexportData(A351_E1_dept_nm))				& """" & vbCr
	IF E2_conf_fg(A351_E2_a_conf_fg) = "U" Then
		Response.Write " 	.frm1.txtTempGlNo.value	= """ & ConvSPChars(UCase(Trim(E2_conf_fg(A351_E2_a_temp_gl_no)))) & """" & vbCr	
	Elseif E2_conf_fg(A351_E2_a_conf_fg) = "C"	Then
		Response.Write " 	.frm1.txtGlNo.value		= """ & ConvSPChars(UCase(Trim(E2_conf_fg(A351_E2_a_gl_no)))) & """" & vbCr	
	End If
	Response.Write " 	.frm1.txtDesc.Value			= """ & ConvSPChars(iexportData(A351_E1_a_temp_gl_temp_gl_desc)) & """" & vbCr		
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	            " & vbCr
	Response.Write " 	.ggoSpread.SSShowData """ & iStrData & """,""F""" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1 & " ,.C_DOC_CUR1,.C_XCH_RATE1    ,""D"",""I"",""X"",""X"")"	& vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1 & " ,.C_DOC_CUR1,.C_OPEN_DOC_AMT1,""A"",""I"",""X"",""X"")"	& vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData4	             " & vbCr
	Response.Write " 	.ggoSpread.SSShowData """ & iStrData1 & """,""F""" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & -1 & "," & -1 & " ,.C_DocCur_2,.C_ExchRate_2,""D"",""I"",""X"",""X"")"	& vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & -1 & "," & -1 & " ,.C_DocCur_2,.C_ItemAmt_2 ,""A"",""I"",""X"",""X"")"	& vbCr	
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
	
	Const A352_I2_a_temp_gl_temp_gl_no = 0
    Const A352_I2_a_temp_gl_temp_gl_dt = 1
    Const A352_I2_a_temp_gl_org_change_id = 2
    Const A352_I2_a_temp_gl_dept_cd = 3
    Const A352_I2_a_temp_gl_gl_type = 4
    Const A352_I2_a_temp_gl_gl_input_type = 5
    Const A352_I2_a_temp_gl_temp_gl_desc = 6
    Const A352_I2_a_temp_gl_project_no = 7

	Dim PAUG035_cAMngClsCardSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim I3_a_cls_no
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetClsNo
	
	Dim iLngMaxRow
	Dim iLngMaxRow3
	Dim iLngRow
	Dim iArrTemp1
	Dim iArrTemp2
	
	'--------------------------------------------------------------------
	'A_GL에 대한 정보  Setting
	'--------------------------------------------------------------------
	iCommandSent = "CREATE"

	I1_b_currency = gCurrency

    ReDim I2_a_gl(7)
	I2_a_gl(A352_I2_a_temp_gl_temp_gl_no)		= ""
	I2_a_gl(A352_I2_a_temp_gl_temp_gl_dt)		= UNIConvDate(Request("txtClsDt"))
	I2_a_gl(A352_I2_a_temp_gl_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_a_gl(A352_I2_a_temp_gl_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A352_I2_a_temp_gl_gl_type)			= Trim(Request("cboGlType")) 
	I2_a_gl(A352_I2_a_temp_gl_gl_input_type)	= Trim(Request("cboGlInputType"))      
	I2_a_gl(A352_I2_a_temp_gl_temp_gl_desc)		= Request("txtDesc")
	I2_a_gl(A352_I2_a_temp_gl_project_no)       = ""

    Dim I5_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A352_I5_a_data_auth_data_BizAreaCd = 0
    Const A352_I5_a_data_auth_data_internal_cd = 1
    Const A352_I5_a_data_auth_data_sub_internal_cd = 2
    Const A352_I5_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I5_a_data_auth(3)
	I5_a_data_auth(A352_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	'--------------------------------------------------------------------
	'A_GL_ITEM에 대한 정보  Setting
	'--------------------------------------------------------------------
	txtSpread = Request("txtSpread")
	
    '--------------------------------------------------------------------
	'A_GL_DTL에 대한 정보  Setting
	'--------------------------------------------------------------------
	txtSpread3 = Request("txtSpread6")
	
	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------	
	Set PAUG035_cAMngClsCardSvr = Server.CreateObject("PAUG035.cAMngClsCardSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
	    
	iStrRetClsNo = PAUG035_cAMngClsCardSvr.A_MANAGE_CLS_CARD_SVR(gStrGlobalCollection, iCommandSent, I1_b_currency, I2_a_gl, _
										I3_a_cls_no, I3_trans_type, txtSpread,  txtSpread3,I5_a_data_auth) 	

	If CheckSYSTEMError(Err, True) = True Then		
		Response.Write "DDD"
		Set PAUG035_cAMngClsCardSvr = Nothing
		Exit Sub
    End If

    Set PAUG035_cAMngClsCardSvr  = Nothing

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.txtClsNo.value = """ & ConvSPChars(Trim(iStrRetClsNo)) & """" & vbCr
    Response.Write "	.DbSaveOk """ & iStrRetClsNo	&			 """" & vbCr    
    Response.Write " End With											" & vbCr
    Response.Write " </Script>											" & vbCr
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
    
	Dim PAUG035_cAMngClsCardSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim I3_a_cls_no
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetClsNo
	
	iCommandSent = "DELETE"
	
	ReDim I2_a_gl(7)

	If Request("txtGlNo") <> "" Then 	
		I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)			= UCase(Trim(Request("txtGlNo")))
	Else 
		I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)			= UCase(Trim(Request("txtTempGlNo")))
	End If		
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_dt)			= UNIConvDate(Request("txtClsDt"))
	I2_a_gl(A393_I2_a_temp_gl_org_change_id)		= Trim(Request("hOrgChangeId"))
	I2_a_gl(A393_I2_a_temp_gl_dept_cd)				= UCase(Trim(Request("txtDeptCd")))
    I2_a_gl(A393_I2_a_temp_gl_gl_type)				= Trim(Request("cboGlType")) 
	I2_a_gl(A393_I2_a_temp_gl_gl_input_type)		= Trim(Request("txtGlinputType"))      
    I2_a_gl(A393_I2_a_temp_gl_temp_gl_desc)			= ""
    I2_a_gl(A393_I2_a_temp_gl_project_no)			= UCase(Trim(Request("txtClsNo")))
    
    Dim I5_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A352_I5_a_data_auth_data_BizAreaCd = 0
    Const A352_I5_a_data_auth_data_internal_cd = 1
    Const A352_I5_a_data_auth_data_sub_internal_cd = 2
    Const A352_I5_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I5_a_data_auth(3)
	I5_a_data_auth(A352_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A352_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))    

	Set PAUG035_cAMngClsCardSvr = Server.CreateObject("PAUG035.cAMngClsCardSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
    
	iStrRetClsNo = PAUG035_cAMngClsCardSvr.A_MANAGE_CLS_CARD_SVR(gStrGlobalCollection, iCommandSent, I1_b_currency, I2_a_gl, _
										I3_a_cls_no, I3_trans_type, txtSpread,  txtSpread3,I5_a_data_auth) 	
	
	If CheckSYSTEMError(Err, True) = True Then		
		Set PAUG035_cAMngClsCardSvr = Nothing
		Exit Sub
    End If
    
    Set PAUG035_cAMngClsCardSvr  = Nothing

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
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next

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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

%>	
	

