<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%					
'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next
Err.Clear

Call HideStatusWnd			'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "A", "NOCOOKIE", "MB")

Dim PADG035

Dim strMode					'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim StrNextKeyTempGlNo		' ���� ä�ǹ�ȣ 
Dim StrNextKeyTempGlDt		' �߻�ä���� ���� �� 
Dim lgStrPrevKeyTempGlNo	' ���� ä�ǹ�ȣ 
Dim lgStrPrevKeyTempGlNo2	' ���� ä�ǹ�ȣ 
Dim lgStrPrevKeyTempGlDt	' �߻�ä���� ���� �� 
Dim iStrBizAreaCd
Dim iStrCboConfFg
Dim iStrFromTempGlDt
Dim iStrToTempGlDt
Dim iLngMaxRow				' ���� �׸����� �ִ�Row

Dim iLngRow

Const C_SHEETMAXROWS = 30
Const iUID_M9999 = 9999		'���� ��ǥ Ŭ���ÿ� ������ǥ�� ��ȸ�ؿ��� �÷��� 

'##########################################################################
'ma ���� mb�� �Ѿ���� �Ķ���� ���� 
strMode = Trim(Request("txtMode"))	'�� : ���� ���¸� ���� 
lgStrPrevKeyTempGlNo = Request("lgStrPrevKeyTempGlNo")
lgStrPrevKeyTempGlNo2 = UNIConvNum(Request("lgStrPrevKeyTempGlNo2"),0)
lgStrPrevKeyTempGlDt = UNIConvDate(Request("lgStrPrevKeyTempGlDt"))

iStrBizAreaCd = Trim(Request("txtBizAreaCd"))
iStrCboConfFg = Trim(Request("cboConfFg"))
iStrFromTempGlDt = UNIConvDate(Request("txtFromTempGlDt"))
istrToTempGlDt = UNIConvDate(Request("txtToTempGlDt"))
iLngMaxRow = UNIConvNum(Request("txtMaxRows"),0)
'###########################################################################

Select Case strMode
Case CStr(UID_M0001)	
	Call SubBizQuery()
Case CStr(UID_M0002)	
	Call SubBizSave()
Case Cstr(iUID_M9999)	
	Call SubBizQuery2()
End Select

Sub SubBizQuery()

	Dim iArrImportView
	Dim iExportView
	Dim iStrData
	Dim iIntLoopCount

	ReDim iArrImportView(5)
	Const C_FromDtATempGlTempGlDt = 0
	Const C_ToDtATempGlTempGlDt = 1
	Const C_BBizAreaBizAreaCd = 2
	Const C_BAcctDeptOrgChangeId = 3
	Const C_ConfFgATempGlConfFg = 4
	Const C_NextTempGlNoATempGlTempGlNo = 5

	'#################################################
	Dim iVarExportView
	Const C_E1_TEMP_GL_NO = 0
	Const C_E1_TEMP_GL_DT = 1
	Const C_E1_ISSUED_DT = 2
	Const C_E1_GL_TYPE = 3
	Const C_E1_GL_INPUT_TYPE = 4
	Const C_E1_CR_AMT = 5
	Const C_E1_CR_LOC_AMT = 6
	Const C_E1_DR_AMT = 7
	Const C_E1_DR_LOC_AMT = 8
	Const C_E1_CONF_FG = 9
	Const C_E1_HQ_BRCH_FG = 10
	Const C_E1_HQ_BRCH_NO = 11
	Const C_E2_GL_NO = 12
	Const C_E2_GL_DT = 13
	Const C_E3_DEPT_CD = 14
	Const C_E3_DEPT_NM = 15
	
	On Error Resume Next
	Err.Clear
	'Import View
	
	If lgStrPrevKeyTempGlNo2 <> 0  And lgStrPrevKeyTempGlNo = "" Then
		Exit Sub
	End If
	
	
	If  lgStrPrevKeyTempGlNo2 = 0 Then
		iArrImportView(C_FromDtATempGlTempGlDt) = iStrFromTempGlDt
		iArrImportView(C_ToDtATempGlTempGlDt) = istrToTempGlDt
		iArrImportView(C_BBizAreaBizAreaCd) = iStrBizAreaCd
		iArrImportView(C_BAcctDeptOrgChangeId) = Request("hOrgChangeID")
		iArrImportView(C_ConfFgATempGlConfFg) = iStrCboConfFg
		iArrImportView(C_NextTempGlNoATempGlTempGlNo) = ""
		lgStrPrevKeyTempGlNo2 = lgStrPrevKeyTempGlNo2 + 1
	Else
		iArrImportView(C_FromDtATempGlTempGlDt) = lgStrPrevKeyTempGlDt
		iArrImportView(C_ToDtATempGlTempGlDt) = istrToTempGlDt
		iArrImportView(C_BBizAreaBizAreaCd) = iStrBizAreaCd
		iArrImportView(C_BAcctDeptOrgChangeId) = Request("hOrgChangeID")
		iArrImportView(C_ConfFgATempGlConfFg) = iStrCboConfFg
		iArrImportView(C_NextTempGlNoATempGlTempGlNo) = lgStrPrevKeyTempGlNo
		lgStrPrevKeyTempGlNo2 = lgStrPrevKeyTempGlNo2 + 1
	End If
	
	 
	
	SET PADG035 = CreateObject("PADG035.cAHqListTmpGlAHqSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If

	Call PADG035.A53018HQ_LIST_TEMP_GL_HQ_SVR(gStrGlobalCollection,C_SHEETMAXROWS,iArrImportView,iVarExportView)

	If CheckSYSTEMError(Err, True) = True Then
		Set PAFG570 = Nothing
		Exit Sub
	End If

		
	If isEmpty(iVarExportView) = False Then
		For iLngRow = 0 To UBound(iVarExportView, 1)
			iIntLoopCount = iIntLoopCount + 1
			If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
				 
				istrData = istrData & Chr(11) & "0"
				istrData = istrData & Chr(11) & " "' ConvSPChars(iVarExportView(iLngRow,C_E1_CONF_FG))
				istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,C_E1_TEMP_GL_DT))
				If Trim(iVarExportView(iLngRow,C_E2_GL_DT)) <> ""  Then
					istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,C_E2_GL_DT))
				Else
					istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,C_E1_TEMP_GL_DT))
				End If

				istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow,C_E1_TEMP_GL_NO))
				istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow,C_E3_DEPT_NM))
				istrData = istrData & Chr(11) & " " 'Currency
				istrData = istrData & Chr(11) & UNINumClientFormat(iVarExportView(iLngRow,C_E1_DR_AMT), ggAmtOfMoney.DecPoint,0)
				istrData = istrData & Chr(11) & UNINumClientFormat(iVarExportView(iLngRow,C_E1_DR_LOC_AMT), ggAmtOfMoney.DecPoint,0)

				istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow,C_E2_GL_NO))
				istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow,C_E1_GL_INPUT_TYPE))
				istrData = istrData & Chr(11) & " " 'InputTypeNm
				istrData = istrData & Chr(11) & ConvSPChars(iVarExportView(iLngRow,C_E1_HQ_BRCH_NO))
				istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
				istrData = istrData & Chr(11) & Chr(12)
			Else
				lgStrPrevKeyTempGlNo = ConvSPChars(iVarExportView(iLngRow,C_E1_TEMP_GL_NO))
				lgStrPrevKeyTempGlDt = UNIDateClientFormat(iVarExportView(iLngRow,C_E1_TEMP_GL_DT))
				Exit For
			End If
		Next
		
		If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					lgStrPrevKeyTempGlNo = ""
					lgStrPrevKeyTempGlDt = ""
		End If
		
		
	End if 
	'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "<Script Language=vbscript>"										 & vbCr
	Response.write "With parent"													 & vbCr
	Response.Write ".ggoSpread.Source = .frm1.vspdData"								 & vbCr
	Response.Write ".ggoSpread.SSShowData """ & istrData & """"						 & vbCr
	Response.Write ".lgStrPrevKeyTempGlDt = """ & lgStrPrevKeyTempGlDt & """"		 & vbCr
	Response.Write ".lgStrPrevKeyTempGlNo = """ & lgStrPrevKeyTempGlNo & """"		 & vbCr
	Response.Write ".lgStrPrevKeyTempGlNo2 = """ & lgStrPrevKeyTempGlNo2 & """"		 & vbCr
	Response.Write ".frm1.hFromTempGlDt.value = """ & Request("FromTempGlDt") & """" & vbCr
	Response.Write ".frm1.hToTempGlDt.value   = """ & Request("ToTempGlDt") & """"	 & vbCr
	Response.Write ".DbQueryOk"														 & vbCr
	Response.write "End With"														 & vbCr
	Response.Write "</Script>"														 & vbCr
	
End Sub

Sub SubBizSave()

	Dim iErrorPosition

    On Error Resume Next
	Err.Clear

	SET PADG035 = CreateObject("PADG035.cAHqCnfmTmpGlSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
	
	CAll PADG035.A53012HQ_CONFIRM_TEMP_GL_SVR(gStrGlobalCollection,Request("txtSpread"),iErrorPosition)

	If CheckSYSTEMError2(Err, True,iErrorPosition & "��","","","","") = True Then
		Set PADG035 = Nothing
		Exit Sub
	End If
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.write "Parent.DbSaveOk"			& vbCr
	Response.Write "</Script>"					& vbCr

End Sub

Sub SubBizQuery2()

	Dim iStrBizAreaCd
	Dim iStrHqBrchNo
	Dim iStrTempGlNo
	Dim iStrChangeOrgID
	Dim iLngMaxLow
	Dim istrData
	Dim iCurRow
	
	'Export View
	Dim iVarExportView
	Const A475_EG1_E1_TEMP_GL_NO = 0
	Const A475_EG1_E1_TEMP_GL_DT = 1
	Const A475_EG1_E1_ISSUED_DT = 2
	Const A475_EG1_E1_GL_TYPE = 3
	Const A475_EG1_E1_GL_INPUT_TYPE = 4
	Const A475_EG1_E1_CR_AMT = 5
	Const A475_EG1_E1_CR_LOC_AMT = 6
	Const A475_EG1_E1_DR_AMT = 7
	Const A475_EG1_E1_DR_LOC_AMT = 8
	Const A475_EG1_E1_CONF_FG = 9
	Const A475_EG1_E1_HQ_BRCH_FG = 10
	Const A475_EG1_E1_HQ_BRCH_NO = 11
	Const A475_EG1_E2_GL_NO = 12
	Const A475_EG1_E2_GL_DT = 13
	Const A475_EG1_E3_DEPT_CD = 14
	Const A475_EG1_E3_DEPT_NM = 15
	
	On Error Resume Next
	Err.Clear

	'MA1�κ��� �Ѿ�� �Ķ���� 
	iStrBizAreaCd = Trim(Request("txtBizAreaCd"))
	iStrHqBrchNo = Trim(Request("txtHqBrchNo"))
	iStrTempGlNo = Trim(Request("txtTempGLNo"))
	iStrChangeOrgID = Trim(Request("hOrgChangeID"))
	iLngMaxLow = Cint(Trim(Request("txtMaxRows")))
	iCurRow	 =		Cint(Request("vspdDataRow"))
	
	SET PADG035 = CreateObject("PADG035.cABrListTmpGlBrchSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If

	Call PADG035.A53018BR_LIST_TEMP_GL_BRCH_SVR(gStrGlobalCollection,C_SHEETMAXROWS,iStrBizAreaCd, iStrHqBrchNo, iStrTempGlNo, iStrChangeOrgID,iVarExportView)

	If CheckSYSTEMError(Err, True) = True Then
		Set PADG035 = Nothing
		Response.Write "Error"
		Exit Sub
	End If

	IF ISEMPTY(iVarExportView) THEN
	ELSE
		For iLngRow = 0 To UBound(iVarExportView, 1)
			istrData = istrData & Chr(11) & "0"
			istrData = istrData & Chr(11) & " " 'iVarExportView(iLngRow,A475_EG1_E1_CONF_FG)
			istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,A475_EG1_E1_TEMP_GL_DT))
			If Trim(iVarExportView(iLngRow,A475_EG1_E2_GL_DT)) <> "" Then
				istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,A475_EG1_E2_GL_DT))
			Else
				istrData = istrData & Chr(11) & UNIDateClientFormat(iVarExportView(iLngRow,A475_EG1_E1_TEMP_GL_DT))
			End If

			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E1_TEMP_GL_NO)
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E3_DEPT_NM)
			istrData = istrData & Chr(11) & " " 'Currency
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E1_DR_AMT)
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E1_DR_LOC_AMT)
			
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E2_GL_NO)
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E1_GL_INPUT_TYPE)
			istrData = istrData & Chr(11) & " " 'InputTypeNm
			istrData = istrData & Chr(11) & iVarExportView(iLngRow,A475_EG1_E1_HQ_BRCH_NO)
			istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
			istrData = istrData & Chr(11) & Chr(12)
		Next
	END IF
'	Response.Write istrData
	'ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "<Script Language=VBscript>"							& vbCr
	Response.write "With parent"										& vbCr
	Response.Write ".ggoSpread.Source = .frm1.vspdData2"				& vbCr
	Response.Write ".ggoSpread.SSShowData """ & istrData &		""""	& vbCr	
	Response.Write " Call .SetVspdData2Checked (" & iCurRow & ")"		& vbCr
	Response.Write ".DbQueryOk2"										& vbCr
	Response.write "End With"											& vbCr
	Response.Write "</Script>"											& vbCr

	Set PADG035 = Nothing

End Sub
%>
