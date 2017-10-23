
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5101mb
'*  4. Program �̸�      : �������� ��� 
'*  5. Program ����      : �������� ��� ���� ���� ��ȸ 
'*  6. Comproxy ����Ʈ   : f5101mb
'*  7. ���� �ۼ������   : 2000/10/12
'*  8. ���� ���������   : 2000/10/12
'*  9. ���� �ۼ���       : ����ȯ 
'* 10. ���� �ۼ���       : Jang Yoon Ki
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                        - 2003/03/23 Oh, Soo Min ���� �����ڵ� (note_acct_cd, 
'*																					rcpt_acct_cd,
'*																					charge_acct_cd �߰�)
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next                                                            '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call HideStatusWnd 

Dim txtNoteNoQry
Dim lgOpModeCRUD
Dim lPtxtNoteNo
Dim sChangeOrgId

sChangeOrgId = GetGlobalInf("gChangeOrgId")
    
                                                              '��: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
'lgErrorStatus     = "NO"
'lgErrorPos        = ""                                                           '��: Set to space
lgOpModeCRUD      = Trim(Request("txtMode"))                                           '��: Read Operation Mode (CRUD)
txtNoteNoQry      = Request("txtNoteNoQry")

Dim strCode																	'�� : Lookup �� �ڵ� ���� ���� 
Dim GroupCount

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'###############single select#################
Const C_NOTE_NO		= 0		'������ȣ 
Const C_NOTE_FG		= 1		'�������� 
Const C_NOTE_STS	= 2		'�������� 
Const C_BP_CD		= 3		'�ŷ�ó 
Const C_BP_NM		= 4		'�ŷ�ó�� 
Const C_BANK_CD		= 5		'���� 
Const C_BANK_NM		= 6		'����� 
Const C_ISSUE_DT	= 7		'������ 
Const C_DUE_DT		= 8		'������ 
Const C_NOTE_AMT	= 9		'�����ݾ� 
Const C_STTL_AMT	= 10	'�����ݾ� 

'############## multi select######################
Const C_STS_DT_SPD			 = 0
Const C_GL_NO_SPD			 = 1
Const C_TEMP_GL_NO_SPD		 = 2
Const C_SEQ_SPD				 = 3
Const C_NOTE_STS_SPD		 = 4
Const C_DC_RATE_SPD			 = 5
Const C_DC_INT_AMT_SPD		 = 6
Const C_CHARGE_AMT_SPD		 = 7
Const C_AMT_SPD				 = 8
Const C_BP_CD_SPD			 = 9
Const C_BP_NM_SPD			 = 10
Const C_BANK_CD_SPD			 = 11
Const C_BANK_NM_SPD			 = 12
Const C_BANK_ACCT_NO_SPD	 = 13
Const C_RCPT_TYPE_SPD		 = 14
Const C_CHG_NOTE_ACCT_CD_SPD = 15
Const C_CHG_NOTE_ACCT_NM_SPD = 16
Const C_NOTE_ACCT_CD_SPD	 = 17
Const C_NOTE_ACCT_NM_SPD     = 18
Const C_DC_INT_ACCT_CD_SPD   = 19
Const C_DC_INT_ACCT_NM_SPD   = 20
Const C_CHARGE_ACCT_CD_SPD   = 21
Const C_CHARGE_ACCT_NM_SPD   = 22
Const C_NOTE_ITEM_DESC       = 23

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '��: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)                                                         '��: Save,Update
         Call SubBizSave()
    Case CStr(UID_M0003)                                                         '��: Delete
         Call SubBizDelete()
End Select

'==================================================================================
'	Name : SubBizQuery()
'	Description : ��ȸ ���� 
'==================================================================================
Sub SubBizQuery()
	On Error Resume Next                                                                 '��: Protect system from crashing
	Err.Clear 

	Dim PAFG535LIST	
	Dim indx
	Dim E1_f_note, EG1_export_group	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount
	Dim I1_f_note
	
	Const A822_I1_note_no = 0    
	Const C_SHEETMAXROWS = 100

	Redim I1_f_note(A822_I1_note_no+4)
	I1_f_note(A822_I1_note_no)   = txtNoteNoQry
	I1_f_note(A822_I1_note_no+1) = lgAuthBizAreaCd
	I1_f_note(A822_I1_note_no+2) = lgInternalCd
	I1_f_note(A822_I1_note_no+3) = lgSubInternalCd
	I1_f_note(A822_I1_note_no+4) = lgAuthUsrID

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
    If Len(Trim(iIntQueryCount))  Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = ""
    End If
    
    Set PAFG535LIST = server.CreateObject ("PAFG535.cFListNoteDtlSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	
    Call PAFG535LIST.FN0028_LIST_NOTE_DTL_SVR(gStrGlobalCollection, I1_f_note, iStrPrevKey, C_SHEETMAXROWS, E1_f_note, EG1_export_group)
	
    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG535LIST = Nothing		
		Exit Sub
    End If

    Set PAFG535LIST = Nothing 
    
	If isempty(E1_f_note) = False Then
		Response.Write "<Script Language=vbscript>  " & vbCr
	   	Response.Write " with parent.frm1           " & vbCr
	   	Response.Write " .txtNoteNo.value	= """ & E1_f_note(C_NOTE_NO) & """													 " & vbCr
	   	Response.Write " .cboNoteFg.value	= """ & E1_f_note(C_NOTE_FG) & """													 " & vbCr
	   	Response.Write " .cboNoteSts.value	= """ & ConvSPChars(E1_f_note(C_NOTE_STS)) & """									 " & vbCr
	   	Response.Write " .txtBpCd.value		= """ & ConvSPChars(E1_f_note(C_BP_CD)) & """										 " & vbCr
	   	Response.Write " .txtBpNM.value		= """ & ConvSPChars(E1_f_note(C_BP_NM)) & """										 " & vbCr
	   	Response.Write " .txtBankCd.Value	= """ & ConvSPChars(E1_f_note(C_BANK_CD)) & """										 " & vbCr
	   	Response.Write " .txtBankNm.Value	= """ & ConvSPChars(E1_f_note(C_BANK_NM)) & """										 " & vbCr
	   	Response.Write " .txtIssueDt.Value	= """ & UNIDateClientFormat(E1_f_note(C_ISSUE_DT)) & """							 " & vbCr
	   	Response.Write " .txtDueDt.Value	= """ & UNIDateClientFormat(E1_f_note(C_DUE_DT)) & """								 " & vbCr
	   	Response.Write " .txtNoteAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_NOTE_AMT),	ggAmtOfMoney.DecPoint	,0) & """" & vbCr
	   	Response.Write " .txtSttlAmt.Text	= """ & UNINumClientFormat(E1_f_note(C_STTL_AMT),	ggAmtOfMoney.DecPoint	,0) & """" & vbCr
		Response.Write "End with				" & vbcr
	    Response.Write "</Script>               " & vbCr
	End If	

	iStrData = ""

	If isempty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group, 1) 	
			iIntLoopCount = iIntLoopCount + 1
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, C_STS_DT_SPD))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_GL_NO_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_TEMP_GL_NO_SPD)))
				iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, C_SEQ_SPD)
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_STS_SPD)))
				iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_DC_RATE_SPD),	ggExchRate.DecPoint		,0)
				iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_DC_INT_AMT_SPD),	ggAmtOfMoney.DecPoint	,0)
				iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_CHARGE_AMT_SPD),	ggAmtOfMoney.DecPoint	,0)
				iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, C_AMT_SPD),		ggAmtOfMoney.DecPoint	,0)
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BP_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BP_NM_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BANK_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BANK_NM_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BANK_ACCT_NO_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_RCPT_TYPE_SPD)))
				iStrData = iStrData & Chr(11) & ""			
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CHG_NOTE_ACCT_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CHG_NOTE_ACCT_NM_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_ACCT_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_ACCT_NM_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_DC_INT_ACCT_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_DC_INT_ACCT_NM_SPD)))				
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CHARGE_ACCT_CD_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CHARGE_ACCT_NM_SPD)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_NOTE_ITEM_DESC)))
				iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
				iStrData = iStrData & Chr(11) & Chr(12)
			Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), C_SEQ_SPD)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
		
	End If

	Response.Write " <Script Language=vbscript>								 " & vbCr
	Response.Write " With parent											 " & vbCr
    Response.Write "	.ggoSpread.Source		= .frm1.vspdData			 " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData	  """ & iStrData		& """" & vbCr
    Response.Write "	.lgPageNo				= """ & iIntQueryCount	& """" & vbCr
    Response.Write "	.lgStrPrevKey			= """ & iStrPrevKey		& """" & vbCr
    Response.Write "	.DbQueryOk											 " & vbCr
    Response.Write "End With												 " & vbCr
    Response.Write "</Script>												 " & vbCr 	
End Sub

'==================================================================================
'	Name : SubBizSave()
'	Description : ����, �ű� 
'==================================================================================
Sub SubBizSave()
	On Error Resume Next                                                                 '��: Protect system from crashing
	Err.Clear 

	Dim PAFG535CU
	Dim lgIntFlgMode

	Dim I1_ief_supplied
	Const C_IEF_SUPPLIED_CUD = 0

	Dim I2_f_note
	Const C_NOTE_NO_CUD = 0

	Dim I3_f_note_item
	Const C_SEQ_CUD				= 0
	Const C_NOTE_STS_CUD		= 1
	Const C_STS_DT_CUD			= 2
	Const C_DC_RATE_CUD			= 3
	Const C_DC_INT_AMT_CUD		= 4
	Const C_CHARGE_AMT_CUD	    = 5
	Const C_STTL_AMT_CUD		= 6
	Const C_GL_NO_CUD			= 7
	Const C_GL_SEQ_CUD			= 8
	Const C_TEMP_GL_NO_CUD	    = 9
	Const C_TEMP_GL_SEQ_CUD	    = 10
	Const C_GL_STS_CUD			= 11
	Const C_BANK_ACCT_NO_CUD    = 12
	Const C_RCPT_TYPE_CUD		= 13
	'2003/03/22 �߰�(�Ա�����, ������ �����ڵ�)
	Const C_NOTE_ACCT_CD_CUD	= 14
	Const C_DC_INT_ACCT_CD_CUD  = 15
	Const C_CHARGE_ACCT_CD_CUD	= 16
	Const C_NOTE_ITEM_DESC      = 17

	Dim I4_b_biz_partner
	Const C_BP_CD_CUD = 0

	Dim I5_b_bank
	Const C_BANK_CD_CUD = 0

	'2003/03/22 �߰�(������ ���� �����ڵ� (DC,DH)
	Dim I6_f_note_item

	Redim I1_ief_supplied(C_IEF_SUPPLIED_CUD)
	Redim I2_f_note(C_NOTE_NO_CUD)
	Redim I3_f_note_item(C_NOTE_ITEM_DESC)
	Redim I4_b_biz_partner(C_BP_CD_CUD)
	Redim I5_b_bank(C_BANK_CD_CUD)

    Dim I7_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
    Const A821_I7_a_data_auth_data_BizAreaCd = 0
    Const A821_I7_a_data_auth_data_internal_cd = 1
    Const A821_I7_a_data_auth_data_sub_internal_cd = 2
    Const A821_I7_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I7_a_data_auth(3)
	I7_a_data_auth(A821_I7_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	I2_f_note = UCase(Trim(Request("txtNoteNoQry")))
		
	I3_f_note_item(C_SEQ_CUD)			 = UNIConvNum(Request("txtSeq"),0)         
	I3_f_note_item(C_NOTE_STS_CUD)		 = UCase(Trim(Request("txtNoteSts1")))
	I3_f_note_item(C_STS_DT_CUD)		 = UNIConvDate(Request("txtStsDt1"))       
	I3_f_note_item(C_DC_RATE_CUD)		 = UNIConvNum(Request("txtDcRate1"),0)     
	I3_f_note_item(C_DC_INT_AMT_CUD)     = UNIConvNum(Request("txtDcIntAmt1"),0)
	I3_f_note_item(C_CHARGE_AMT_CUD)	 = UNIConvNum(Request("txtChargeAmt1"),0)     
	I3_f_note_item(C_STTL_AMT_CUD)		 = UNIConvNum(Request("txtSttlAmt1"),0)
	I3_f_note_item(C_GL_NO_CUD)			 = "" 
	I3_f_note_item(C_GL_SEQ_CUD)		 = 0
	I3_f_note_item(C_TEMP_GL_NO_CUD)	 = "" 
	I3_f_note_item(C_TEMP_GL_SEQ_CUD)	 = 0
	I3_f_note_item(C_GL_STS_CUD)		 = "" 
	I3_f_note_item(C_BANK_ACCT_NO_CUD)	 = UCase(Trim(Request("txtBankAcct1")))
	I3_f_note_item(C_RCPT_TYPE_CUD)		 = UCase(Trim(Request("txtRcptType1")))
	I3_f_note_item(C_NOTE_ACCT_CD_CUD)	 = UCase(Trim(Request("txtNoteAcctCd")))
	I3_f_note_item(C_DC_INT_ACCT_CD_CUD) = UCase(Trim(Request("txtDcIntAcctCd")))
	I3_f_note_item(C_CHARGE_ACCT_CD_CUD) = UCase(Trim(Request("txtChargeAcctCd")))
	I3_f_note_item(C_NOTE_ITEM_DESC)	 = Trim(Request("txtDesc"))
	
	I4_b_biz_partner = UCase(Trim(Request("txtBpCd1")))
	I5_b_bank		 = UCase(Trim(Request("txtBankCd1")))
	I6_f_note_item	 = UCase(Trim(Request("txtChgNoteAcctCd")))
	lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '��: Read Operayion Mode (CREATE, UPDATE)

	Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                             '�� : Create
			I1_ief_supplied(C_IEF_SUPPLIED_CUD) = "C"	
        Case  OPMD_UMODE          
			I1_ief_supplied(C_IEF_SUPPLIED_CUD) = "U"
    End Select		

    Set PAFG535CU = server.CreateObject("PAFG535.cFMngNoteItmSvr")   

    Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                             '�� : Create
			  Call PAFG535CU.FN0022_MANAGE_NOTE_ITEM_SVR(gStrGlobalCollection, _
														"CREATE", _
														sChangeOrgId, _
														I1_ief_supplied, _
														I2_f_note, _
														I3_f_note_item, _
														I4_b_biz_partner, _
														I5_b_bank, _
														I6_f_note_item, _
														I7_a_data_auth)
        
        Case  OPMD_UMODE          
			  Call PAFG535CU.FN0022_MANAGE_NOTE_ITEM_SVR(gStrGlobalCollection, _
														"UPDATE", _
														sChangeOrgId, _
														I1_ief_supplied, _
														I2_f_note, _
														I3_f_note_item, _
														I4_b_biz_partner, _
														I5_b_bank, _
														I6_f_note_item, _
														I7_a_data_auth)
    End Select

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG535CU = nothing
		Exit Sub	
    End If
	 
    Set PAFG535CU = nothing

	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk()							" & vbCr
    Response.Write "</Script>									" & vbCr
End Sub

'==================================================================================
'	Name : SubBizDelete()
'	Description : ���� 
'==================================================================================
Sub SubBizDelete()
	On Error Resume Next                                                                 '��: Protect system from crashing
	
	Dim PAFG535D
	Dim iarrData

	Dim I1_ief_supplied
	Const C_IEF_SUPPLIED_CUD = 0

	Dim I2_f_note

	Dim I3_f_note_item
	Const C_SEQ_CUD				= 0
	Const C_GL_NO_CUD			= 1

	Redim I1_ief_supplied(C_IEF_SUPPLIED_CUD)
	Redim I3_f_note_item(C_GL_NO_CUD)

    Dim I7_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
    Const A821_I7_a_data_auth_data_BizAreaCd = 0
    Const A821_I7_a_data_auth_data_internal_cd = 1
    Const A821_I7_a_data_auth_data_sub_internal_cd = 2
    Const A821_I7_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I7_a_data_auth(3)
	I7_a_data_auth(A821_I7_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I7_a_data_auth(A821_I7_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	I1_ief_supplied(C_IEF_SUPPLIED_CUD) = "D"
	
	I2_f_note		= UCase(Trim(Request("txtNoteNo")))
	
	I3_f_note_item(C_SEQ_CUD)		= UNIConvNum(Request("txtSeq"),0)         
	I3_f_note_item(C_GL_NO_CUD)		= UCase(Trim(Request("txtGlNo1")))	
	
     Set PAFG535D = server.CreateObject ("PAFG535.cFMngNoteItmSvr")     
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
    Call PAFG535D.FN0022_MANAGE_NOTE_ITEM_SVR(gStrGlobalCollection, _
											"DELETE", _
											sChangeOrgId, _
											I1_ief_supplied, _
											I2_f_note, _
											I3_f_note_item, _
											,_
											,_
											,_
											I7_a_data_auth)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG535D = nothing
		Exit Sub
    End If
	 
    Set PAFG535D = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub
%>	
