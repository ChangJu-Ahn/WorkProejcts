<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5114ma
'*  4. Program �̸�      : ����ī��ó�� 
'*  5. Program ����      : ����ī��ó�� 
'*  6. Comproxy ����Ʈ   : f5114ma
'*  7. ���� �ۼ������   : 2002/06/19
'*  8. ���� ���������   : 2002/08/09
'*  9. ���� �ۼ���       : 
'* 10. ���� �ۼ���       : Shin Myoung_Ha
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           : 1. UniConvNum()���� - 2002/08/02
'*						   2. FilterVar() ����(Com���� ����) - 2002/08/08
'*						   3. ��¥, ���� OCX�� TEXT�� VALUE�� �߸��Ȼ�� ���� - 2002/08/09
'*                         
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

On Error Resume Next			' ��: 
ERR.CLEAR

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim LngMaxRow					' ���� �׸����� �ִ�Row
Dim LngRow

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim lgStrPrevKeyNoteNo			' Note NO ���� �� 
Dim lgStrPrevKeyGlNo
Dim iStrProcFg
Dim iStrData

Const C_SHEETMAXROWS = 100

'Import View ��� 
Const A707_I1_PROGFG = 0
Const A707_I1_DUEDTEND = 1
Const A707_I1_STSDTSTART = 2
Const A707_I1_STSDTEND = 3
Const A707_I1_BPCD = 4
Const A707_I1_BANKCD = 5
Const A707_I1_PREVKEYNOTENO = 6
Const A707_I1_PREVKEYGINO = 7

'EXPORTS Group View ��� 
Const A707_EG1_E1_bp_cd = 0
Const A707_EG1_E1_bp_nm = 1
Const A707_EG1_E1_gl_no = 2
Const A707_EG1_E1_gl_dt = 3
Const A707_EG1_E1_minor_cd = 4
Const A707_EG1_E1_minor_nm = 5
Const A707_EG1_E1_note_no = 6
Const A707_EG1_E1_note_amt = 7
Const A707_EG1_E1_issue_dt = 8
Const A707_EG1_E1_due_dt = 9
Const A707_EG1_E1_note_sts = 10
Const A707_EG1_E1_dept_cd = 11
Const A707_EG1_E1_dept_nm = 12
Const A707_EG1_E1_org_change_id = 13
Const A707_EG1_E1_bank_cd = 14
Const A707_EG1_E1_bank_nm = 15
Const A707_EG1_E1_temp_gl_no = 16
Const A707_EG1_E1_temp_gl_dt = 17
Const A707_EG1_E1_rcpt_type = 18

strMode = Trim(Request("txtMode"))			'�� : ���� ���¸� ���� 
LngMaxRow = Request("txtMaxRows")			'��: Read Operation Mode (CRUD)
iStrProcFg = Trim(Request("cboProcFg"))			'ó��(CG), ���(DG) ���� 

lgStrPrevKeyNoteNo = "" & UCase(Trim(Request("lgStrPrevKeyNoteNo")))
lgStrPrevKeyGlNo = "" & UCase(Trim(Request("lgStrPrevKeyGlNo")))

'Response.Write gStrGlobalCollection
'Response.Write iStrProcFg
'Response.Write strMode & "<br>"
'Response.Write UID_M0002 & "<br>"

Select Case Trim(strMode) 
	Case Trim(UID_M0001)
		Call SubBizQuery()
	Case Trim(UID_M0002)
		Call SubBizSave()		
End Select

Sub SubBizQuery()		
	Dim PAFG570
	Dim iArrImportView
	Dim iVarExportView
	Dim iIntLoopCount	
	Dim iLngRow
	
	Redim iArrImportView(8)
	ON ERROR RESUME NEXT
	ERR.CLEAR
	
	'IMPORTVIEW SETTING
	iArrImportView(A707_I1_PROGFG) = Trim(Request("cboProcFg"))
	iArrImportView(A707_I1_DUEDTEND) = Trim(Request("txtDueDtEnd"))
	iArrImportView(A707_I1_STSDTSTART) = UNIConvDate(Trim(Request("txtStsDtStart")))
	iArrImportView(A707_I1_STSDTEND) = UNIConvDate(Trim(Request("txtStsDtEnd")))
	iArrImportView(A707_I1_BPCD) = Trim(Request("txtBpCd"))
	iArrImportView(A707_I1_BANKCD) = Trim(Request("txtBankCd"))
	iArrImportView(A707_I1_PREVKEYNOTENO) = Trim(Request("lgStrPrevKeyNoteNo"))
	iArrImportView(A707_I1_PREVKEYGINO) = Trim(Request("lgStrPrevKeyGlNo"))

		
	
	SET PAFG570 = CreateObject("PAFG570.cFListCardForBtchSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub	
	End If
	
	Call PAFG570.FC0048_LIST_CARD_FOR_BATCH_SVR(gStrGlobalCollection,C_SHEETMAXROWS,iArrImportView,iVarExportView)	
			
	If CheckSYSTEMError(Err, True) = True Then
		Set PAFG570 = Nothing		
		Exit Sub
	End If
	
	Select Case UCase(LTrim(RTrim(iStrProcFg)))
	Case Trim("CG")	
			
			For iLngRow = 0 To UBound(iVarExportView, 1)
				'iIntLoopCount = iIntLoopCount + 1
				If  iLngRow < C_SHEETMAXROWS Then			
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_note_no))				
					iStrData = iStrData & Chr(11) & UNINumClientFormat(iVarExportView(iLngRow, A707_EG1_E1_note_amt), ggAmtOfMoney.DecPoint, 0)
					iStrData = iStrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow, A707_EG1_E1_due_dt))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bp_cd))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bp_nm))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bank_cd))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bank_nm))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_dept_cd))
					iStrData = iStrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_dept_nm))
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & LngMaxRow + iLngRow + 1
					iStrData = iStrData & Chr(11) & Chr(12)
				Else			
					lgStrPrevKeyNoteNo = iVarExportView(iLngRow, A707_EG1_E1_note_no)
					lgStrPrevKeyGlNo = ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_gl_no))					
				End If
			Next
		
	Case "DG"	
		
			For iLngRow = 0 To UBound(iVarExportView, 1)				
				if iLngRow < C_SHEETMAXROWS Then				
					istrData = istrData & Chr(11) & 0
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_note_no))			'note_no
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_temp_gl_no))				'temp_glno
					istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow, A707_EG1_E1_temp_gl_dt))		'temp_gldt
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_gl_no))				'glno
					istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow, A707_EG1_E1_gl_dt))		'gldt											  
					istrData = istrData & Chr(11) & UNINumClientFormat(iVarExportView(iLngRow, A707_EG1_E1_note_amt), ggAmtOfMoney.DecPoint, 0)	'noteamt
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bp_cd))				'bpcd
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bp_nm))				'bpnm
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bank_cd))			'bankcd
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_bank_nm))			'banknm
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_dept_cd))			'deptcd
					istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_dept_nm))			'deptnm				
					istrData = istrData & Chr(11) & LngMaxRow + iLngRow +1
					istrData = istrData & Chr(11) & Chr(12)				
				Else				
					lgStrPrevKeyNoteNo = ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_note_no))
					lgStrPrevKeyGlNo = ConvSPChars(iVarExportView(iLngRow, A707_EG1_E1_gl_no))				
				End If
				
			Next
		
	
	End Select

	if iLngRow <= C_SHEETMAXROWS then
		lgStrPrevKeyNoteNo = ""
		lgStrPrevKeyGlNo = ""
	end if
	
	'Response.Write lgStrPrevKeyNoteNo
	test = "test"
	'ȭ�鿡 ����Ÿ ���� 
	Select Case RTrim(LTrim(iStrProcFg))
	Case Trim("CG")		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "With parent" & vbCr
		Response.Write ".frm1.vspdData.Redraw = False" & vbCr
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCr
		Response.Write ".ggoSpread.SSShowData """ & istrData & """" & vbCr
		Response.Write ".lgStrPrevKeyNoteNo = """ & lgStrPrevKeyNoteNo & """" & vbCr
		Response.Write ".frm1.vspdData.Redraw = True" & vbCr
		Response.Write ".frm1.hProcFg.value = """ & iArrImportView(A707_I1_PROGFG) & """" & vbCr		
		Response.Write ".frm1.hDueDtEnd.value = """ & iArrImportView(A707_I1_DUEDTEND) & """" & vbCr		
		Response.Write ".frm1.hBpCd.value = """ & ConvSPChars(iArrImportView(A707_I1_BPCD)) & """" & vbCr
		Response.Write ".frm1.hBankCd.Value = """ & ConvSPChars(iArrImportView(A707_I1_BANKCD)) & """" & vbCr		
		Response.Write ".DbQueryOK" & vbCr
		Response.write "End With" & vbCr		
		Response.Write "</script>" & vbCr						
	Case Trim("DG")
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "With parent" & vbCr
		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCr 
		Response.Write ".ggoSpread.SSShowData """ & iStrData & """" & vbCr
		Response.Write ".lgStrPrevKeyNoteNo = """ & lgStrPrevKeyNoteNo & """" & vbCr
		Response.Write ".lgStrPrevKeyGlNo   = """ & lgStrPrevKeyGlNo & """" & vbCr		
		Response.Write ".frm1.hProcFg.value	= """ & iArrImportView(A707_I1_PROGFG) & """" & vbCr
		Response.Write ".frm1.hStsDtStart.value	= """ & iArrImportView(A707_I1_STSDTSTART) & """" & vbCr
		Response.Write ".frm1.hStsDtEnd.value = """ & iArrImportView(A707_I1_STSDTEND) & """" & vbCr			
		Response.Write ".frm1.hBpCd.value = """ & iArrImportView(A707_I1_BPCD) & """" & vbCr
		Response.Write ".frm1.hBankCd.value = """ & iArrImportView(A707_I1_BANKCD) & """" & vbCr
		Response.Write ".DbQueryOK" & vbCr		 		
		Response.write "End With" & vbCr		
		Response.Write "</script>" & vbCr	
	End Select
	
	
	
End Sub

'==================================================================================
'	Name : SubBizSaveMuliti()
'	Description : ��Ƽ���� ���� 
'==================================================================================
Sub SubBizSave()
	
	Dim iArrImportView
	Const ORG_CHINGE_ID = 0
	Const DEPT_ID = 1
	Const RCPT_TYPE = 2
	Const RCPT_ACCT_CD = 3
	Const CHARGE = 4
	Const CHARGE_ACCT_CD = 5
	Const BANK_CD = 6
	Const BANK_ACC_NO = 7
	Const GL_DT = 8
	
	Dim PAFG570
	
	ReDim iArrImportView(8)	
	
	ON ERROR RESUME NEXT
	ERR.CLEAR	
		
	iStrProcFg = Trim(Request("cboProcFg"))			'ó��(CG), ���(DG) ���� 

    Dim I1_a_data_auth								
    Const A814_I1_a_data_auth_data_BizAreaCd = 0
    Const A814_I1_a_data_auth_data_internal_cd = 1
    Const A814_I1_a_data_auth_data_sub_internal_cd = 2
    Const A814_I1_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A814_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A814_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A814_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A814_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	Select Case Trim(Request("hProcFg"))
		Case "CG"		
			iArrImportView(0) = Request("hOrgChangeId")
			iArrImportView(1) = Trim(Request("txtDeptCd"))
			iArrImportView(2) = Trim(Request("txtRcptType"))
			iArrImportView(3) = Trim(Request("txtNoteAcctCd"))
			iArrImportView(4) = CDbl((UniConvNum(Request("txtChargeAmt"),1)))
			iArrImportView(5) = Trim(Request("txtChargeAcctCd"))
			iArrImportView(6) = Trim(Request("txtBankCd"))
			iArrImportView(7) = Trim(Request("txtBankAcctNo"))
			iArrImportView(8) = UNIConvDate(Request("txtGLDt"))

			Set PAFG570 = CreateObject("PAFG570.cFBtchCardSvr")
			
			If CheckSYSTEMError(Err,True) = True Then
				Exit Sub
			End If
			
			CAll PAFG570.F_BATCH_CARD_SVR(gStrGlobalCollection,_
										Request("txtSpread"), _
										iArrImportView, _
										I1_a_data_auth)
			
			If CheckSYSTEMError(Err, True) = True Then
				Set PAFG570 = Nothing			
				Exit Sub
			End If
		Case "DG"
			SET PAFG570 = CreateObject("PAFG570.cFBtchCardSvr")
	
			If CheckSYSTEMError(Err, True) = True Then
				Exit Sub	
			End If
					
			Call PAFG570.F_BATCH_CARD_SVR(gStrGlobalCollection,_
										  Request("txtSpread"), _
										  , _
										  I1_a_data_auth)
			
			If CheckSYSTEMError(Err, True) = True Then
				Set PAFG570 = Nothing
				Exit Sub
			End If
	End Select 
		
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DbSaveOk" & vbCr
	Response.Write "</Script>" & vbCr
End Sub
%>
