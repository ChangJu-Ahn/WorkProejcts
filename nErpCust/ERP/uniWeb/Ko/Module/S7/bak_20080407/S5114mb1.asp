<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5114MA1
'*  4. Program Name         : ����ä�Ǽ��ݳ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : S51119LookupBillHdrSvr, S51158ListBillCollectingSvr, S51151MaintBillCollectingSvr
'*							  S51115PostOpenArSvr, Fr0019LookupPrrcptSvr
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Dim strOpModeCRUD
    
	Const lsPOSTFLAG	= "PostFlag"									

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I","*", "NOCOOKIE", "MB")	
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
 
    strOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

    Select Case strOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
        Case CStr(lsPOSTFLAG)        ' ����ä��Ȯ��ó�� 
             Call SubBizPostFlag()
    End Select
'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()

	Dim lgArrGlFlag
	Dim lgStrGlFlag
	Dim lgStrPostFlag
	Dim lgStrGlNo
	Dim lgCurrency

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 
                                                      '��: Protect system from crashing
    '-----------------------
    ' ��������� �о�´�.
    '-----------------------
    Call SubOpenDB(lgObjConn)
    call SubMakeSQLStatements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	    lgObjRs.Close
	    lgObjConn.Close
	    Set lgObjRs = Nothing
	    Set lgObjConn = Nothing
		'����ä�������� �����ϴ�.
		Call DisplayMsgBox("205100", vbInformation, "", "", I_MKSCRIPT)             '��: No data is found. 

		Response.Write "<Script Language=vbscript>"		& vbCr
		Response.Write "parent.SetDefaultVal"           & vbCr
		Response.Write "Call parent.SetToolBar(""11000000000011"")"           & vbCr
		Response.Write "</Script>"						& vbCr

	    Exit Sub
	End If

	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "With parent.frm1"					& vbCr


		'##### Rounding Logic #####
		'�׻� �ŷ�ȭ�� �켱 
	lgCurrency = ConvSPChars(lgObjRs("Cur"))
			
	Response.Write ".txtCurrency.value		= """		& lgCurrency          & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX "		& vbCr 		

	'-----------------------
	' ��������� ������ ǥ���Ѵ�.
	'----------------------- 

	'���ܸ��⿩�� 
	If UCase(Trim(lgObjRs("except_flag"))) = "Y" Then	
	    Response.Write ".rdoExceptBillYes.checked = True	"   & vbCr	
	Elseif UCase(Trim(lgObjRs("except_flag"))) = "N" Then		
	    Response.Write ".rdoExceptBillNo.checked = True		"   & vbCr	
	End If

	'Ȯ������ 
	If UCase(Trim(lgObjRs("post_flag"))) = "Y" Then	
	    Response.Write ".rdoPostYes.checked = True	"    & vbCr	
	Elseif UCase(Trim(lgObjRs("post_flag"))) = "N" Then		
	    Response.Write ".rdoPostNo.checked = True	"    & vbCr	
	End If

		'�ֹ�ó 
	Response.Write ".txtSoldtoParty.value   = """ & ConvSPChars(lgObjRs("Sold_to_Party"))		& """" & vbCr		
	Response.Write ".txtSoldtoPartyNm.value	= """ & ConvSPChars(lgObjRs("Sold_to_Party_Nm"))	& """" & vbCr		
		'����ä���� 
	Response.Write ".txtBillDt.Text			= """ & UNIDateClientFormat(lgObjRs("bill_dt"))		& """" & vbCr
		'����ä������ 
	Response.Write ".txtBillTypeCd.value    = """ & ConvSPChars(lgObjRs("bill_type"))			& """" & vbCr
	Response.Write ".txtBillTypeNm.value    = """ & ConvSPChars(lgObjRs("bill_type_nm"))		& """" & vbCr
		'�����׷� 
	Response.Write ".txtSalesGrpCd.value    = """ & ConvSPChars(lgObjRs("sales_grp"))			& """" & vbCr
	Response.Write ".txtSalesGrpNm.value    = """ & ConvSPChars(lgObjRs("sales_grp_nm"))		& """" & vbCr

		'�Ѹ���ä�Ǳݾ� 
    Response.Write ".txtTotBillAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(CDbl(lgObjRs("bill_amt")) + CDbl(lgObjRs("vat_amt")) + CDbl(lgObjRs("deposit_amt")), lgCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
		'�Ѹ���ä���ڱ��ݾ� 
    Response.Write ".txtTotBillAmtLoc.Text	= """ & UNIConvNumDBToCompanyByCurrency(CDbl(lgObjRs("bill_amt_loc")) + CDbl(lgObjRs("vat_amt_loc")) + CDbl(lgObjRs("deposit_amt_loc")), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
'kek
	Response.Write ".txtLocCur.value		=  UCase(parent.parent.gCurrency)		"							&   vbCr
		'�� ���ݾ� 
    Response.Write ".txtSumBillAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr
		'�� �����ڱ��� 
    Response.Write ".txtSumLocBillAmt.Text  = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt_loc"), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr

		'ȯ�� 
	Response.Write ".HXchRate.value			= """ & UNINumClientFormat(lgObjRs("Xchg_Rate"), ggExchRate.DecPoint, 0)      & """" & vbCr		
		'ȯ�������� 
	Response.Write ".HXchRateOp.value       = """ & lgObjRs("xchg_rate_op")						& """" & vbCr		

		'����ä��������� 
	Response.Write ".txtSts.value           = """ & ConvSPChars(lgObjRs("sts"))                 & """" & vbCr		

	Response.Write ".txtHBillNo.value		= """ & ConvSPChars(Request("txtConBillNo"))		& """" & vbCr		


	lgStrPostFlag = lgObjRs("post_flag")
	lgStrGlNo = Trim(lgObjRs("gl_no"))
	If lgStrPostFlag = "Y" AND Len(lgStrGlNo) Then
		lgArrGlFlag = Split(lgStrGlNo, Chr(11)) 
		lgStrGlFlag = lgArrGlFlag(0)

		If lgArrGlFlag(0) = "G" Then	
			'ȸ����ǥ��ȣ 
			Response.Write ".txtGLNo.value       = """ & lgArrGlFlag(1)            & """" & vbCr			 
		ElseIf lgArrGlFlag(0) = "T" Then
			'������ǥ��ȣ 
			Response.Write ".txtTempGLNo.value   = """ & lgArrGlFlag(1)            & """" & vbCr		
		Else
			'Batch��ȣ 
			Response.Write ".txtBatchNo.value    = """ & lgArrGlFlag(1)            & """" & vbCr	
		End If
	Else
		Response.Write ".txtGLNo.value       = """""     & vbCr	
		Response.Write ".txtTempGLNo.value   = """""     & vbCr	
		Response.Write ".txtBatchNo.value    = """""     & vbCr	
				
	End If

	If lgStrPostFlag = "Y" Then
		Response.Write ".btnPostFlag.value = ""Ȯ�����"""      & vbCr
		if lgStrGlFlag = "G" Or lgStrGlFlag = "T" Then
			Response.Write ".btnGLView.disabled = False"            & vbCr
		Else
			Response.Write ".btnGLView.disabled = True"             & vbCr
		End if
	Else
		Response.Write ".btnPostFlag.value	= ""Ȯ��"""         & vbCr
		Response.Write ".btnGLView.disabled = True"                 & vbCr
	End If

	Response.Write ".txtHPostFlag.value    = """ & lgStrPostFlag	& """" & vbCr	

	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing


'	If UCASE(gCurrency) = UCASE(lgCurrency) Then
'		Response.Write ".vspdData.Col = parent.C_BillLocAmt		:	.vspdData.ColHidden = True"          & vbCr		
'		Response.Write ".vspdData.Col = parent.C_XchRate		:	.vspdData.ColHidden = True"          & vbCr		
'		Response.Write ".vspdData.Col = parent.C_XchCalop		:	.vspdData.ColHidden = True"          & vbCr				
'	Else
'		Response.Write ".vspdData.Col = parent.C_BillLocAmt		:	.vspdData.ColHidden = False"         & vbCr				
'		Response.Write ".vspdData.Col = parent.C_XchRate		:	.vspdData.ColHidden = False"         & vbCr				
'		Response.Write ".vspdData.Col = parent.C_XchCalop		:	.vspdData.ColHidden = False"         & vbCr								
'	End If

	If UCase(gCurrency) = UCase(lgCurrency) Then
		Response.Write "parent.SetInitSpreadSheet	"                   & vbCr
	Else
		Response.Write "parent.SetInitSpreadSheet2	"                   & vbCr	
	End If


		'-----------------------
		' Rounding Column Set
		'----------------------- 
	Response.Write "parent.CurFormatNumSprSheet	"					& vbCr
	Response.Write "parent.DbQueryOk	"                           & vbCr					'��: ��ȸ�� ���� 
    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
										
' kek : SubBizQueryMulti

	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
    Dim iStrNextKey  	

    Dim pS7G158
    
	Dim iarrValue    
   
    Dim I1_s_bill_collecting_collect_seq 
    Dim I2_s_bill_hdr_bill_no 
    Dim EG1_exp_grp 

    Const C_SHEETMAXROWS_D  = 100

    '[CONVERSION INFORMATION]  IMPORTS View ��� 

    'Const S566_I1_collect_seq = 0    '[CONVERSION INFORMATION]  View Name : imp_next s_bill_collecting

    '[CONVERSION INFORMATION]  IMPORTS View ��� 

    'Const S566_I2_bill_no = 0		  '[CONVERSION INFORMATION]  View Name : imp s_bill_hdr

    '[CONVERSION INFORMATION]  EXPORTS Group View ��� 
    '[CONVERSION INFORMATION] ===========================================================================
    '[CONVERSION INFORMATION]  Group Name : exp_grp
    Const S566_EG1_E1_minor_nm = 0		 '[CONVERSION INFORMATION]  View Name : exp_item_collect_type_nm b_minor
    Const S566_EG1_E2_bank_nm = 1		 '[CONVERSION INFORMATION]  View Name : exp_item b_bank
    Const S566_EG1_E3_collect_seq = 2    '[CONVERSION INFORMATION]  View Name : exp_item s_bill_collecting
    Const S566_EG1_E3_collect_type = 3
    Const S566_EG1_E3_collect_doc_amt = 4
    Const S566_EG1_E3_collect_loc_amt = 5
    Const S566_EG1_E3_note_no = 6
    Const S566_EG1_E3_bank_cd = 7
    Const S566_EG1_E3_xch_rate = 8
    Const S566_EG1_E3_xch_rate_op = 9
    Const S566_EG1_E3_pre_rcpt_no = 10
    Const S566_EG1_E3_bank_acct_no = 11
    Const S566_EG1_E3_ext1_qty = 12
    Const S566_EG1_E3_ext2_qty = 13
    Const S566_EG1_E3_ext3_qty = 14
    Const S566_EG1_E3_ext1_amt = 15
    Const S566_EG1_E3_ext2_amt = 16
    Const S566_EG1_E3_ext3_amt = 17
    Const S566_EG1_E3_ext1_cd = 18
    Const S566_EG1_E3_ext2_cd = 19
    Const S566_EG1_E3_ext3_cd = 20
    Const S566_EG1_E3_remark = 21
    '[CONVERSION INFORMATION] ===========================================================================

    '[CONVERSION INFORMATION]  EXPORTS View ��� 

    'Const S566_E1_collect_seq = 0    '[CONVERSION INFORMATION]  View Name : exp_next s_bill_collecting

    On Error Resume Next
    Err.Clear 

	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                  '��: Next Key	

	If iStrPrevKey <> "" then					
		iarrValue = Split(iStrPrevKey, gColSep)
		I1_s_bill_collecting_collect_seq = cdbl(Trim(iarrValue(0)))
	else			
		I1_s_bill_collecting_collect_seq = 0
	End If	
	
    If Request("txtConBillNo") = "" Then										'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("��ȸ ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If
	
	I2_s_bill_hdr_bill_no  = Trim(Request("txtConBillNo"))
	
    Set pS7G158 = Server.CreateObject("PS7G158.cSLtBlCollectingSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  

	Call pS7G158.S_LIST_BILL_COLLECTING_SVR(gStrGlobalCollection, Cint(C_SHEETMAXROWS_D), _
										I1_s_bill_collecting_collect_seq, I2_s_bill_hdr_bill_no, EG1_exp_grp )        
    
	If CheckSYSTEMError(Err,True) = True Then
       Set pS7G158 = Nothing
		Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.frm1.txtConBillNo.focus  " & vbCr   		
		Response.Write "</Script>      " & vbCr      
       Exit Sub
    End If  
    
    Set pS7G158 = Nothing

    iLngMaxRow  = CLng(Request("txtMaxRows"))										 '��: Fetechd Count      

	For iLngRow = 0 To UBound(EG1_exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_collect_seq)) 
           Exit For
        End If 

			 '�������� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_collect_type)) 			 
			 '���������˾� 
		istrData = istrData & Chr(11) & ""	 
			 '���������� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E1_minor_nm)) 			 
			 '���ݾ� 
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow, S566_EG1_E3_collect_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")  
			 '�����ڱ��ݾ� 
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow, S566_EG1_E3_collect_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  
			 '���� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_bank_cd)) 			 
			 '�����˾� 
		istrData = istrData & Chr(11) & ""	 
			 '����� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E2_bank_nm)) 			 			 
			'������¹�ȣ 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_bank_acct_no)) 			 			 
			'������¹�ȣ�˾� 
		istrData = istrData & Chr(11) & ""	 
			'������ȣ 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_note_no)) 			 			 			
			'������ȣ�˾� 
		istrData = istrData & Chr(11) & ""	 
			'�����ݹ�ȣ 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_pre_rcpt_no)) 			 			 			
			'�����ݹ�ȣ�˾� 
		istrData = istrData & Chr(11) & ""	 
			'��� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_remark)) 			 			 			
			'ȯ�� 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S566_EG1_E3_xch_rate), ggExchRate.DecPoint, 0)
			'ȯ�������� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S566_EG1_E3_xch_rate_op)) 			 			 			
			'���ݼ��� 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S566_EG1_E3_collect_seq), 0, 0)			
       
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12) 

    Next    

    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent	       " & vbCr

    Response.Write "   .ggoSpread.Source          = .frm1.vspdData		        " & vbCr
    Response.Write "   .ggoSpread.SSShowDataByClip        """ & istrData	       & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey	   & """" & vbCr

    Response.Write "   .DbQueryOk  " & vbCr   
    
	If lgStrPostFlag = "Y" Then
		Response.Write "   .SetPostYesSpreadColor(1)  " & vbCr   	
	ELSE
	    Response.Write "   .SetQuerySpreadColor(1)  " & vbCr   
	End If
    
    Response.Write " End With       " & vbCr															    	
    Response.Write "</Script>      " & vbCr      


End Sub	
    
'============================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()


End Sub
'============================================
' Name : SubBizSave
' Desc : Save Data 
'============================================
Sub SubBizSave()
    
End Sub
'============================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti()        
	Dim pS7G151 
	Dim iCommandSent
	Dim iErrorPosition

	Dim I1_s_bill_hdr_bill_no
	Dim I2_s_wks_user_user_id

    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear																			 '��: Clear Error status                                                            

	I1_s_bill_hdr_bill_no = UCase(Trim(Request("txtHBillNo")))
	'I2_s_wks_user_user_id = Trim(Request("txtInsrtUserId"))
	
	Set pS7G151 = Server.CreateObject("PS7G151.cSBillCollectingSvr")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    Call pS7G151.S_MAINT_BILL_COLLECTING_SVR(gStrGlobalCollection, Trim(Request("txtSpread")), _
                           I1_s_bill_hdr_bill_no, I2_s_wks_user_user_id, iErrorPosition)
											      
    If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
       Set pS7G151 = Nothing
       Exit Sub
	End If
	
    Set pS7G151 = Nothing

    Response.Write "<Script language=vbs> " & vbCr            
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr  		                                                                    
    
End Sub    
'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================
' Name : SubBizPostFlag
' Desc : PostFlag
'============================================
Sub SubBizPostFlag()
    
    Dim pS7G115
    Dim itxtFlgMode
	Dim pvCB
	
	Dim I1_s_bill_hdr_bill_no
                
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
	I1_s_bill_hdr_bill_no = Trim(Request("txtHBillNo"))
	pvCB = "F"
    
    Set pS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               
	call pS7G115.S_POST_OPEN_AR_SVR(pvCB, gStrGlobalCollection, I1_s_bill_hdr_bill_no)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pS7G115 = Nothing
		Exit Sub
	End If     
    '-----------------------
	'Result data display area
	'----------------------- 
	Set pS7G115 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "		& vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo ( " & FilterVar(Request("txtConBillNo"), "''", "S") & ", " & FilterVar("%", "''", "S") & ", " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Q", "''", "S") & " , " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
