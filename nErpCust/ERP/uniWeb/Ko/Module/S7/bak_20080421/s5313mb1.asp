<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5313MB1
'*  4. Program Name         : ���ݰ�꼭��ȣ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G338.cSListTaxDocNoSvr,PS7G331.cSTaxDocNoSvr
'*  7. Modified date(First) : 2001/06/26
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2001/06/26 : 6�� ȭ�� layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'*							  -2002/11/14 : UI���� ����	
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call LoadBasisGlobalInf()

Dim lgOpModeCRUD

On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")

Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
		'Call SubBizQuery()
		Call SubBizQueryMulti()
	Case CStr(UID_M0002)
		'Call SubBizSave()
		Call SubBizSaveMulti()
	 Case CStr(UID_M0003)                                                         '��: Delete
        'Call SubBizDelete()
     Case CStr("PostFlag")																'��: ���� ��û 
		Call SubPostFlag()
End Select

'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================
' Name : SubBizSave
' Desc : Save Data 
'============================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()

	Dim iS7G338												'�� : ���ݰ�꼭��ȣ��� �Է�/����/������ ComProxy Dll ��� ���� 
	Dim StrNextKey											' ���� �� 
	Dim ILngMaxRow											' ���� �׸����� �ִ�Row
	Dim ILngRow
	Dim istrData
	
	Dim I1_s_tax_doc_no
	Dim I2_s_tax_doc_no
	Dim EG1_exp_grp
	Dim E1_s_tax_doc_no
	
	Const C_SHEETMAXROWS_D = 100
	
	Const I1_tax_doc_no = 0									'I1_s_tax_doc_no
    Const I1_tax_book_no = 1
    Const I1_used_flag = 2
    Const I1_usage_flag = 3
    
    Const E1_tax_doc_no = 0									'EG1_exp_grp
    Const E1_tax_book_no = 1
    Const E1_tax_book_seq = 2
    Const E1_used_flag = 3
    Const E1_usage_flag = 4
    Const E1_insrt_user_id = 5
    Const E1_insrt_dt = 6
    Const E1_updt_user_id = 7
    Const E1_updt_dt = 8
    Const E1_ext1_qty = 9
    Const E1_ext2_qty = 10
    Const E1_ext3_qty = 11
    Const E1_ext1_amt = 12
    Const E1_ext2_amt = 13
    Const E1_ext3_amt = 14
    Const E1_ext1_cd = 15
    Const E1_ext2_cd = 16
    Const E1_ext3_cd = 17
    Const E1_expiry_date = 18
    Const E1_created_meth = 19
    Const E1_tax_bill_no = 20
    Const E1_issued_dt = 21
    
    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear	
    
    Redim I1_s_tax_doc_no(3)
	'-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_s_tax_doc_no(I1_tax_doc_no) = Trim(Request("txtTaxDocBillNo"))
    I1_s_tax_doc_no(I1_tax_book_no) = UNIConvNum(Request("txtBookNo"),0)
    I1_s_tax_doc_no(I1_used_flag) = Trim(Request("HUsed"))
    I1_s_tax_doc_no(I1_usage_flag) = Trim(Request("HUseFlag"))
	I2_s_tax_doc_no = Trim(Request("lgStrPrevKey"))
	
	'-----------------------
    ' ���ݰ�꼭��ȣ�� �о�´�.
    '-----------------------
    Set iS7G338 = Server.CreateObject("PS7G338.cSListTaxDocNoSvr")    
 
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   
    
    Call iS7G338.S_LIST_TAX_DOC_NO_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
											I1_s_tax_doc_no, I2_s_tax_doc_no,_
											EG1_exp_grp, E1_s_tax_doc_no) 
    
    If CheckSYSTEMError(Err,True) = True Then
		Set iS7G338 = Nothing
		Response.Write "<Script Language=vbscript>"			& vbCr
        Response.Write "Parent.frm1.txtTaxDocBillNo.focus"		& vbCr    
        Response.Write "</Script>"							& vbCr
		Exit Sub
    End If   
            
	Set iS7G338 = Nothing	
    
    ILngMaxRow  = CLng(Request("txtMaxRows"))										'��: Fetechd Count     
	
	'-----------------------
	'Result data display area
	'----------------------- 
	For ILngRow = 0 To UBound(EG1_exp_grp, 1)
		If  ILngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_exp_grp(ILngRow, E1_tax_doc_no))
		   Exit For
        End If  
	
    	'-----------------------
		' ���ݰ�꼭��ȣ�� ������ ǥ���Ѵ�.
		'----------------------- 
		 '���ݰ�꼭��ȣ 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, E1_tax_doc_no))
		'å��ȣ(��) 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(ILngRow, E1_tax_book_no), 0, 0)
		'å��ȣ(ȣ) 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(ILngRow, E1_tax_book_seq), 0, 0)
		'��뿩�� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, E1_usage_flag))
		
		'��ȿ�� ���� 
		if EG1_exp_grp(ILngRow, E1_expiry_date) = "2999-12-31" then
			istrData = istrData & Chr(11) & ""
		else
			istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(ILngRow, E1_expiry_date))
		end if
									
		'������� �߰� 
		Select Case	ConvSPChars(EG1_exp_grp(ILngRow, E1_created_meth))
			Case "A"
			istrData = istrData & Chr(11) & "Auto"
			Case "M"
			istrData = istrData & Chr(11) & "Manual"
			Case "P"
			istrData = istrData & Chr(11) & "Pre-fixed"
			Case "X"
			istrData = istrData & Chr(11) & "Mixed"
		End Select

		'������ 
		Select Case	ConvSPChars(EG1_exp_grp(ILngRow, E1_used_flag))
			Case "C"
			istrData = istrData & Chr(11) & "Created"
			Case "R"
			istrData = istrData & Chr(11) & "Referenced"
			Case "I"
			istrData = istrData & Chr(11) & "Issued"
			Case "D"
			istrData = istrData & Chr(11) & "Deleted"
		End Select
		
		'���ݰ�꼭������ȣ ���� 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, E1_tax_bill_no))
			
		'������ ���� 
		If EG1_exp_grp(ILngRow, E1_issued_dt) = "2999-12-31" Then
			istrData = istrData & Chr(11) & ""
		Else
			istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(ILngRow, E1_issued_dt))
		End If

		istrData = istrData & Chr(11) & ILngMaxRow + ILngRow 
		istrData = istrData & Chr(11) & Chr(12)

	Next
	
	Response.Write "<Script language=vbs> "												& vbCr
	Response.Write "With parent"														& vbCr   
    Response.Write " .frm1.HTaxBillDocNo.value	= """ & Request("ConvSPChars(txtTaxDocBillNo)")		& """" & vbCr
    Response.Write " .frm1.HBookNo.value		= """ & Request("txtBookNo")			& """" & vbCr    
    Response.Write " .frm1.HUsed.value			= """ & Request("HUsed")				& """" & vbCr        
    Response.Write " .ggoSpread.Source	=         .frm1.vspdData"	    	& vbCr    
    Response.Write " .ggoSpread.SSShowDataByClip		  """ & istrData			& """" & vbCr   
    Response.Write " .lgStrPrevKey				= """ & StrNextKey						& """" & vbCr    
    Response.Write " .DbQueryOk "														& vbCr   
    Response.Write "End With "															& vbCr   
    Response.Write "</Script> "		
	Response.End																				'��: Process End

End Sub

'============================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti() 

	Dim iS7G331												'�� : ���ݰ�꼭��ȣ��� ��ȸ�� ComProxy Dll ��� ���� 
	Dim iErrorPosition
	Dim itxtSpread
	
	On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear	
    									
    Set iS7G331 = Server.CreateObject("PS7G331.cSTaxDocNoSvr")  
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    itxtSpread = Trim(Request("txtSpread"))
    call iS7G331.S_MAINT_TAX_DOC_NO_SVR(gStrGlobalCollection, itxtSpread, iErrorPosition)
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
       Set iS7G331 = Nothing
       Exit Sub
	End If
	
	Set iS7G331 = Nothing
	
    Response.Write "<Script Language=vbscript> "	& vbCr         
    Response.Write " Parent.DBSaveOk "				& vbCr   
    Response.Write "</Script> "           
	Response.End																				'��: Process End
    
End Sub
%>
