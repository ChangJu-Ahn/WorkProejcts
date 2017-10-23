
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : ������ǥ��� 
'*  3. Program ID        : a5101mb
'*  4. Program �̸�      : ������ǥ ��� 
'*  5. Program ����      : ������ǥ������ ���, ����, ����, ��ȸ 
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2000/09/22
'*  8. ���� ���������   : 2001/02/12
'*  9. ���� �ۼ���       : �ۺ��� 
'* 10. ���� �ۼ���       : ������ 
'* 11. ��ü comment      :
'*
'********************************************************************************************** 
													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

%>

<%
     
    Dim lgOpModeCRUD
'    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
            'Call SubBizDelete()
             Call SubBizDeleteMulti()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Const C_AGl_gl_no = 0
	Const C_AGlItem_item_seq = 1
	
	Const A327_E1_a_gl_gl_no = 0
	Const A327_E1_a_gl_gl_dt = 1
	Const A327_E1_a_gl_gl_type = 2
	Const A327_E1_a_gl_input_type = 3
	Const A327_E1_a_gl_cr_amt = 4
	Const A327_E1_a_gl_cr_loc_amt = 5
	Const A327_E1_a_gl_dr_amt = 6
	Const A327_E1_a_gl_dr_loc_amt = 7
	Const A327_E1_a_gl_gl_desc = 8
	Const A327_E1_a_gl_project_no = 9
	Const A327_E1_a_gl_org_change_id = 10
	Const A327_E1_a_gl_dept_cd = 11
	Const A327_E1_a_gl_dept_nm = 12
	
	Const A327_EG1_a_gl_item_item_seq		= 0
	Const A327_EG1_a_gl_item_dept_cd		= 1
	Const A327_EG1_a_gl_item_dept_nm		= 2
	Const A327_EG1_a_gl_item_acct_cd		= 3
	Const A327_EG1_a_gl_item_acct_nm		= 4
	Const A327_EG1_a_gl_item_dr_cr_fg		= 5	
	Const A327_EG1_a_gl_item_item_amt		= 6
	Const A327_EG1_a_gl_item_item_loc_amt	= 7
	Const A327_EG1_a_gl_item_vat_type		= 8
	Const A327_EG1_a_gl_item_item_desc		= 9
	Const A327_EG1_a_gl_item_xch_rate		= 10
	Const A327_EG1_a_gl_item_doc_cur		= 11
	Const A327_EG1_a_gl_item_project_no		= 12
	Const A327_EG1_a_gl_item_gl_no			= 13
	Const A327_EG1_a_gl_item_org_change_id	= 14
	Const A327_EG1_a_gl_item_acct_type		= 15

	Dim PAGG020_cALkUpGlSvr
	Dim iStrData
    Dim iexportData
    Dim iexportData1
    Dim iLngRow
    Dim iLngCol
    Dim iStrCurrency
    
    Dim iArrKey
    ReDim iArrKey(5)
    
    Dim zDataAuth

	On Error Resume Next
    Err.Clear	
	
'==================================================================================================================================	
'���Ѱ����� ���� �߰��� �κ� 
'==================================================================================================================================
   
  if Request("lgAuthorityFlag") = "Y" then
  	  Set zDataAuth = Server.CreateObject("DataAuthorityCheck.CheckMethod")    
      
      If Err.Number <> 0 Then
  			Set zDataAuth = Nothing												        '��: ComProxy Unload
  			Response.End														        '��: �����Ͻ� ���� ó���� ������ 
  	  End If    
    	zDataAuth.importUsrId = gUsrID
    	zDataAuth.importModuleCd = "A"
   	 	zDataAuth.importCheckFactor(1) = "DT"
    	zDataAuth.importCodeValue(1) = UCase(Trim(ConvSPChars(pA51019.ExportBAcctDeptDeptCd)))
    	zDataAuth.ImportOpMode(1) = "S"
    	zDataAuth.importConnectionString = gADODBConnString
    	zDataAuth.Execute
      If Err.Number <> 0 Then
  			Set zDataAuth = Nothing												        '��: ComProxy Unload
  			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)					
  			Response.End														        '��: �����Ͻ� ���� ó���� ������ 
  		End If    
  
  		If Not (zDataAuth.OperationStatusMessage = MSG_OK_STR) Then
	  	    Select Case zDataAuth.OperationStatusMessage  	    
              Case "216001"
                  Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, zDataAuth.OperationErrorValue, zDataAuth.OperationErrorAuthority, I_MKSCRIPT)
              Case Else
                  Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
              End Select
  			Set zDataAuth = Nothing
  			Response.End 
  		End If			
	  
		For LngRow = 1 To GroupCount	
			zDataAuth.importUsrId = gUsrID
			zDataAuth.importModuleCd = "A"
   		 	zDataAuth.importCheckFactor(1) = "DT"
			zDataAuth.importCodeValue(1) = UCase(Trim(ConvSPChars(pA51019.OutGrpBAcctDeptDeptCd(LngRow))))
			zDataAuth.ImportOpMode(1) = "S"
			zDataAuth.importConnectionString = gADODBConnString
			zDataAuth.Execute
		  If Err.Number <> 0 Then
  				Set zDataAuth = Nothing												        '��: ComProxy Unload
  				Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)					
  				Response.End														        '��: �����Ͻ� ���� ó���� ������ 
  			End If    
  
  			If Not (zDataAuth.OperationStatusMessage = MSG_OK_STR) Then
		  	    Select Case zDataAuth.OperationStatusMessage  	    
		          Case "216001"
		              Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, zDataAuth.OperationErrorValue, zDataAuth.OperationErrorAuthority, I_MKSCRIPT)
		          Case Else
		              Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		          End Select
  				Set zDataAuth = Nothing
  				Response.End 
  			End If			
		next	  
	 Set zDataAuth = Nothing
  end if
'==================================================================================================================================	  
    
    iArrKey(C_AGl_gl_no) = UCase(Trim(Request("txtGlNo")))
    iArrKey(C_AGlItem_item_seq) = "0"

	' ���Ѱ��� �߰� 
    iArrKey(C_AGlItem_item_seq+1) = lgAuthBizAreaCd
    iArrKey(C_AGlItem_item_seq+2) = lgInternalCd
    iArrKey(C_AGlItem_item_seq+3) = lgSubInternalCd
    iArrKey(C_AGlItem_item_seq+4) = lgAuthUsrID

    Set PAGG020_cALkUpGlSvr = Server.CreateObject("PAGG020.cALkUpGlSvr")
    
    If CheckSYSTEMError(Err, True) = True Then
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write " With Parent " & vbCr				
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"")"	       & vbCr
		Response.Write " End With " & vbCr
		Response.Write "</Script> " & vbCr
		Exit Sub
    End If
    
	Call PAGG020_cALkUpGlSvr.A_LOOKUP_GL_SVR(gStrGlobalCollection, iArrKey, iexportData, iexportData1)
	
	If CheckSYSTEMError(Err, True) = True Then		
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write " With Parent " & vbCr				
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"")"	       & vbCr
		Response.Write " End With " & vbCr
		Response.Write "</Script> " & vbCr			
		Set PAGG020_cALkUpGlSvr = Nothing
		Exit Sub
    End If    

    Set PAGG020_cALkUpGlSvr = Nothing
    
	iStrCurrency = ConvSPChars(iexportData1(0, A327_EG1_a_gl_item_doc_cur))
    iStrData = ""	

	For iLngRow = 0 To UBound(iexportData1, 1) 
		iStrCurrency = ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_doc_cur))

		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_item_seq))
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iexportData1(iLngRow, A327_EG1_a_gl_item_dept_cd)))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_dept_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_acct_cd))
		iStrData = iStrData & Chr(11) & ""		
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_acct_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(UCase(iexportData1(iLngRow, A327_EG1_a_gl_item_dr_cr_fg)))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_doc_cur))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & UNINumClientFormat(iexportData1(iLngRow, A327_EG1_a_gl_item_xch_rate),	ggExchRate.DecPoint,	 0)' UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_xch_rate),iStrCurrency,ggExchRateNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A327_EG1_a_gl_item_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A327_EG1_a_gl_item_item_loc_amt),gCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_item_desc))
		iStrData = iStrData & Chr(11) & ConvSPChars(UCase(Trim(iexportData1(iLngRow, A327_EG1_a_gl_item_vat_type))))		
		iStrData = iStrData & Chr(11) & ""

		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A327_EG1_a_gl_item_acct_cd))
		iStrData = iStrData & Chr(11) & iLngRow + 1
		
		iStrData = iStrData & Chr(11) & Chr(12)
	Next

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent " & vbCr		
	Response.Write " 	.frm1.txtGlNo.value			= """ & ConvSPChars(UCase(Trim(iexportData(A327_E1_a_gl_gl_no))))			& """" & vbCr
	Response.Write "	.frm1.txtGLDt.Text			= """ & UNIDateClientFormat(iexportData(A327_E1_a_gl_gl_dt))				& """" & vbCr
	Response.Write " 	.frm1.cboGLtype.value		= """ & ConvSPChars(iexportData(A327_E1_a_gl_gl_type))						& """" & vbCr
	Response.Write " 	.frm1.txtGlinputType.value	= """ & ConvSPChars(iexportData(A327_E1_a_gl_input_type))					& """" & vbCr
	Response.Write " 	.frm1.cboGlInputType.Value	= """ & ConvSPChars(iexportData(A327_E1_a_gl_input_type))					& """" & vbCr
'	Response.Write " 	.frm1.txtCrAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A327_E1_a_gl_cr_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr
	Response.Write " 	.frm1.txtCrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A327_E1_a_gl_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
'	Response.Write " 	.frm1.txtDrAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A327_E1_a_gl_dr_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr 
	Response.Write " 	.frm1.txtDrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A327_E1_a_gl_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr							
	Response.Write " 	.frm1.txtDesc.Value			= """ & ConvSPChars(iexportData(A327_E1_a_gl_gl_desc))						& """" & vbCr
	Response.Write " 	.frm1.txtDeptCd.value		= """ & ConvSPChars(UCase(Trim(iexportData(A327_E1_a_gl_dept_cd))))			& """" & vbCr
	Response.Write " 	.frm1.txtDeptNm.value		= """ & ConvSPChars(iexportData(A327_E1_a_gl_dept_nm))						& """" & vbCr
'	Response.Write " 	.frm1.txtDocCur.value		= """ & ConvSPChars(iStrCurrency)											& """" & vbCr			
	Response.Write " 	.frm1.htxtGlNo.value		= """ & ConvSPChars(UCase(Trim(iexportData(A327_E1_a_gl_gl_no))))			& """" & vbCr			
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	      " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData   & """ ,""F""" & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & 1 & "," & iLngRow + 1 & ",.C_DocCur,.C_ItemAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & 1 & "," & iLngRow + 1 & ",.C_DocCur,.C_Exchrate,  ""D"" ,""I"",""X"",""X"")" & vbCr	
	Response.Write " 	.DbQueryOk								  " & vbCr
	Response.Write " End With " & vbCr
	Response.Write "</Script> " & vbCr 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Const A381_I2_a_gl_gl_no = 0
    Const A381_I2_a_gl_gl_dt = 1
    Const A381_I2_a_gl_org_change_id = 2
    Const A381_I2_a_gl_dept_cd = 3
    Const A381_I2_a_gl_gl_type = 4
    Const A381_I2_a_gl_gl_input_type = 5
    Const A381_I2_a_gl_gl_desc = 6
    Const A381_I2_a_gl_project_no = 7

	Dim PAGG020_cAMngGlSvr
	Dim pvCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetGlNo

	Dim I5_a_data_auth
	Const A381_I5_a_data_auth_data_BizAreaCd = 0
	Const A381_I5_a_data_auth_data_internal_cd = 1
	Const A381_I5_a_data_auth_data_sub_internal_cd = 2
	Const A381_I5_a_data_auth_data_auth_usr_id = 3
	
	Dim iLngMaxRow
	Dim iLngMaxRow3
	Dim iLngRow
	Dim iArrTemp1
	Dim iArrTemp2
	
	Dim zDataAuth
	
	On Error Resume Next
	Err.Clear
		
'	'==================================================================================================================================	
'	'���Ѱ����� ���� �߰��� �κ� 
'	'==================================================================================================================================
'   
'	If Request("txtAuthorityFlag") = "Y" then
'  	  Set zDataAuth = Server.CreateObject("DataAuthorityCheck.CheckMethod")    
'      
'      If Err.Number <> 0 Then
'  			Set zDataAuth = Nothing												        '��: ComProxy Unload
'  			Response.End														        '��: �����Ͻ� ���� ó���� ������ 
'  		End If    
'    	zDataAuth.importUsrId = gUsrID
'    	zDataAuth.importModuleCd = "A"
'   	 	zDataAuth.importCheckFactor(1) = "DT"
'    	zDataAuth.importCodeValue(1) = UCase(Trim(Request("txtDeptCd")))
'    	zDataAuth.ImportOpMode(1) = "I"
'    	zDataAuth.importConnectionString = gADODBConnString
'    	zDataAuth.Execute
'      If Err.Number <> 0 Then
'  			Set zDataAuth = Nothing												        '��: ComProxy Unload
'  			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)					
'  			Response.End														        '��: �����Ͻ� ���� ó���� ������ 
'  		End If    
'  
'  		If Not (zDataAuth.OperationStatusMessage = MSG_OK_STR) Then
'	  	    Select Case zDataAuth.OperationStatusMessage  	    
'              Case "216001"
'                  Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, zDataAuth.OperationErrorValue, zDataAuth.OperationErrorAuthority, I_MKSCRIPT)
'              Case Else
'                  Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
'              End Select
'  			Set zDataAuth = Nothing
'  			Response.End 
'  		End If			
'  	  Set zDataAuth = Nothing
'	  end if
'==================================================================================================================================		
	
	iLngMaxRow  = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 
    iLngMaxRow3 = CInt(Request("txtMaxRows3"))
	
	Redim I5_a_data_auth(3)
	I5_a_data_auth(A381_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))	
	
	'--------------------------------------------------------------------
	'A_GL�� ���� ����  Setting
	'--------------------------------------------------------------------
	pvCommandSent = Request("txtCommandMode")	'Spread Sheet ������ ��� �ִ� Element�� 
	I1_b_currency = gCurrency
    
    ReDim I2_a_gl(7)
	I2_a_gl(A381_I2_a_gl_gl_no)			= UCase(Trim(Request("txtGlNo")))
	I2_a_gl(A381_I2_a_gl_gl_dt)			= UNIConvDate(Request("txtGlDt"))
	I2_a_gl(A381_I2_a_gl_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_a_gl(A381_I2_a_gl_dept_cd)		= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A381_I2_a_gl_gl_type)       = Trim(Request("cboGlType")) 
	I2_a_gl(A381_I2_a_gl_gl_input_type) = Trim(Request("cboGlInputType"))      
	I2_a_gl(A381_I2_a_gl_gl_desc)		= Request("txtDesc")

	'--------------------------------------------------------------------
	'A_GL_ITEM�� ���� ����  Setting
	'--------------------------------------------------------------------
    iArrTemp1 = Split(Request("txtSpread"), gRowSep)	'ITEM SPREAD	
 
   	For iLngRow = 1 To iLngMaxRow
		iArrTemp2 = Split(iArrTemp1(iLngRow-1), gColSep)
        txtSpread = txtSpread & "C" & gColSep
		' ��ü ������ ������ ����(Update, Insert) �� ���� 
		If iArrTemp2(0) <> "D" Then	
'==================================================================================================================================	
'���Ѱ����� ���� �߰��� �κ� 
'==================================================================================================================================    
'			if Request("txtAuthorityFlag") = "Y" then
'						
'				zDataAuth.importUsrId = gUsrID
'    			zDataAuth.importModuleCd = "A"
'   	 			zDataAuth.importCheckFactor(1) = "DT"
'    			zDataAuth.importCodeValue(1) = UCase(Trim(arrVal(3)))
'    			zDataAuth.ImportOpMode(1) = "I"
'    			zDataAuth.importConnectionString = gADODBConnString
'    			zDataAuth.Execute
'				If Err.Number <> 0 Then
'  					Set zDataAuth = Nothing												        '��: ComProxy Unload
'  					Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)					
'  					Response.End														        '��: �����Ͻ� ���� ó���� ������ 
'  				End If    
'  
'  				If Not (zDataAuth.OperationStatusMessage = MSG_OK_STR) Then
'  					Select Case zDataAuth.OperationStatusMessage  	    
'						Case "216001"
'							Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, zDataAuth.OperationErrorValue, zDataAuth.OperationErrorAuthority, I_MKSCRIPT)
'						Case Else
'							Call DisplayMsgBox(zDataAuth.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
'			        End Select
'  					Set zDataAuth = Nothing
'  					Response.End 
'  				End If			  			
'			end if
'==================================================================================================================================
			
		    txtSpread = txtSpread & Cint(iArrTemp2(1))										& gColSep	'current row
		    txtSpread = txtSpread & Cint(iArrTemp2(2))										& gColSep	'ItemSEQ 
			txtSpread = txtSpread & iArrTemp2(4)											& gColSep	'�����ڵ� 
			txtSpread = txtSpread & iArrTemp2(5)											& gColSep	'���뱸�� 
			txtSpread = txtSpread & Request("hOrgChangeId")									& gColSep			
			txtSpread = txtSpread & iArrTemp2(3)											& gColSep	'�μ�	
			txtSpread = txtSpread & iArrTemp2(11)											& gColSep   '�ŷ���ȭ 
			If Trim(iArrTemp2(9)) = "" then								'ȯ�� 
				txtSpread = txtSpread & "0"													& gColSep	
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(9))									& gColSep	
			End if
			
			txtSpread = txtSpread & iArrTemp2(10)											& gColSep				
     	    
     	    If Trim(iArrTemp2(6)) = "" then																'�ŷ��ݾ� 
				txtSpread = txtSpread & "0"													& gColSep	
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(6))									& gColSep	
			End if			
			If Trim(iArrTemp2(7)) = "" then																'�ڱ��ݾ� 
				txtSpread = txtSpread & ""													& gColSep					
			Else
				txtSpread = txtSpread & CDbl(iArrTemp2(7))									& gColSep	
			End if
			txtSpread = txtSpread & iArrTemp2(8)											& gRowSep	'��� 
		End If
	Next

    '--------------------------------------------------------------------
	'A_GL_DTL�� ���� ����  Setting
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
	'�����ϱ�.
	'--------------------------------------------------------------------	
	Set PAGG020_cAMngGlSvr = Server.CreateObject("PAGG020.cAMngGlSvr")	
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iStrRetGlNo = PAGG020_cAMngGlSvr.A_MANAGE_GL_SVR(gStrGlobalCollection, pvCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3, ,gDsnNo,I5_a_data_auth)

	If CheckSYSTEMError(Err, True) = True Then		
		Set PAGG020_cAMngGlSvr = Nothing
		Exit Sub
    End If
    
    Set PAGG020_cAMngGlSvr  = Nothing    

	Response.Write " <Script Language=vbscript>											" & vbCr
	Response.Write " With parent												        " & vbCr
    Response.Write "	.DbSaveOk """ & ConvSPChars(iStrRetGlNo)	&				 """" & vbCr    
    Response.Write " End With															" & vbCr
    Response.Write " </Script>															" & vbCr
End Sub    


'============================================================================================================
' Name : SubBizDeleteMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDeleteMulti()
	Const A381_I2_a_gl_gl_no = 0
    Const A381_I2_a_gl_gl_dt = 1
    Const A381_I2_a_gl_org_change_id = 2
    Const A381_I2_a_gl_dept_cd = 3
    Const A381_I2_a_gl_gl_type = 4
    Const A381_I2_a_gl_gl_input_type = 5
    Const A381_I2_a_gl_gl_desc = 6
    Const A381_I2_a_gl_project_no = 7

	Dim PAGG020_cAMngGlSvr
	Dim pvCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetGlNo
	
	Dim I5_a_data_auth
	Const A381_I5_a_data_auth_data_BizAreaCd = 0
	Const A381_I5_a_data_auth_data_internal_cd = 1
	Const A381_I5_a_data_auth_data_sub_internal_cd = 2
	Const A381_I5_a_data_auth_data_auth_usr_id = 3	
	
	On Error Resume Next
	Err.Clear
	
	pvCommandSent = "DELETE"
	ReDim I2_a_gl(7)

	I2_a_gl(A381_I2_a_gl_gl_no)			= UCase(Trim(Request("txtGlNo")))
	I2_a_gl(A381_I2_a_gl_gl_dt)			= UNIConvDate(Request("txtGlDt"))
	I2_a_gl(A381_I2_a_gl_org_change_id)	= Trim(Request("txtOrgChangeId"))
	
	I2_a_gl(A381_I2_a_gl_dept_cd)		= UCase(Trim(Request("txtDeptCd")))
   'I2_a_gl(A381_I2_a_gl_gl_type)        = Trim(Request("cboGlType")) 
	I2_a_gl(A381_I2_a_gl_gl_input_type)  = Trim(Request("txtGlinputType"))      
   'I2_a_gl(A381_I2_a_gl_gl_desc)		= Request("txtDesc") 

	Redim I5_a_data_auth(3)
	I5_a_data_auth(A381_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A381_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	Set PAGG020_cAMngGlSvr = Server.CreateObject("PAGG020.cAMngGlSvr")	

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
    
	iStrRetGlNo = PAGG020_cAMngGlSvr.A_MANAGE_GL_SVR(gStrGlobalCollection, pvCommandSent, "", I2_a_gl, "", "", ,gDsnNo,I5_a_data_auth)
	
	If CheckSYSTEMError(Err, True) = True Then		
		Set PAGG020_cAMngGlSvr = Nothing
		Exit Sub
    End If
    
    Set PAGG020_cAMngGlSvr  = Nothing
	
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

%>
