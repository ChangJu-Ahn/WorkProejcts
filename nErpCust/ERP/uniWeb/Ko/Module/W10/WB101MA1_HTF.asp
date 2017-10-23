<%
'======================================================================================================
'*  1. Function Name        : �Ű��� ���� 
'*  3. Program ID           : WB101MA1
'*  4. Program Name         : WB101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcCompanyInfo

Set lgcCompanyInfo = Nothing	' -- �ʱ�ȭ 

Class C_COMPANY_HISTORY
	' -- ���̺��� �÷����� 
	Dim TAX_DOC_CD
	Dim CO_NM
	Dim CO_ADDR
	Dim OWN_RGST_NO
	Dim LAW_RGST_NO
	Dim REPRE_NM
	Dim REPRE_RGST_NO
	Dim TEL_NO
	Dim COMP_TYPE1
	Dim DEBT_MULTIPLE
	Dim COMP_TYPE2
	Dim TAX_OFFICE
	Dim HOLDING_COMP_FLG
	Dim IND_CLASS
	Dim IND_TYPE
	Dim FOUNDATION_DT
	Dim HOME_TAX_USR_ID
	Dim HOME_TAX_E_MAIL
	Dim HOME_TAX_MAIN_IND
	Dim FISC_START_DT
	Dim FISC_END_DT
	Dim HOME_ANY_START_DT
	Dim HOME_ANY_END_DT
	Dim BANK_CD
	Dim BANK_BRANCH
	Dim BANK_DPST
	Dim BANK_ACCT_NO
	Dim INCOM_DT
	Dim HOME_FILE_MAKE_DT
	Dim REVISION_YM
	Dim EX_RECON_FLG
	Dim EX_54_FLG
	Dim AGENT_NM
	Dim RECON_BAN_NO
	Dim RECON_MGT_NO
	Dim AGENT_TEL_NO
	Dim AGENT_RGST_NO
	Dim REQUEST_DT
	Dim APPO_NO
	Dim APPO_DT
	Dim APPO_DESC
	Dim File_Name
	
	Dim lgoRs1	' -- �����ݱ׸� 
	
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs1
				 
		On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		LoadData = False
		 
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		CO_NM			= oRs1("CO_NM")
		CO_ADDR			= oRs1("CO_ADDR")
		OWN_RGST_NO		= oRs1("OWN_RGST_NO")
		LAW_RGST_NO		= oRs1("LAW_RGST_NO")
		REPRE_NM			= oRs1("REPRE_NM")
		REPRE_RGST_NO	= oRs1("REPRE_RGST_NO")
		TEL_NO			= oRs1("TEL_NO")
		COMP_TYPE1		= oRs1("COMP_TYPE1")
		DEBT_MULTIPLE	= oRs1("DEBT_MULTIPLE")
		COMP_TYPE2		= oRs1("COMP_TYPE2")
		TAX_OFFICE		= oRs1("TAX_OFFICE")
		HOLDING_COMP_FLG	= oRs1("HOLDING_COMP_FLG")
		IND_CLASS		= oRs1("IND_CLASS")
		IND_TYPE			= oRs1("IND_TYPE")
		FOUNDATION_DT	= oRs1("FOUNDATION_DT")
		HOME_TAX_USR_ID	= oRs1("HOME_TAX_USR_ID")
		HOME_TAX_E_MAIL	= oRs1("HOME_TAX_E_MAIL")
		HOME_TAX_MAIN_IND= oRs1("HOME_TAX_MAIN_IND")
		FISC_START_DT	= oRs1("FISC_START_DT")
		FISC_END_DT		= oRs1("FISC_END_DT")
		HOME_ANY_START_DT= oRs1("HOME_ANY_START_DT")
		HOME_ANY_END_DT	= oRs1("HOME_ANY_END_DT")
		BANK_CD			= oRs1("BANK_CD")
		BANK_BRANCH		= oRs1("BANK_BRANCH")
		BANK_DPST		= oRs1("BANK_DPST")
		BANK_ACCT_NO		= oRs1("BANK_ACCT_NO")
		INCOM_DT			= oRs1("INCOM_DT")
		HOME_FILE_MAKE_DT= oRs1("HOME_FILE_MAKE_DT")
		REVISION_YM		= oRs1("REVISION_YM")
		EX_RECON_FLG		= oRs1("EX_RECON_FLG")
		EX_54_FLG		= oRs1("EX_54_FLG")
		AGENT_NM			= oRs1("AGENT_NM")
		RECON_BAN_NO		= oRs1("RECON_BAN_NO")
		RECON_MGT_NO		= oRs1("RECON_MGT_NO")
		AGENT_TEL_NO		= oRs1("AGENT_TEL_NO")
		AGENT_RGST_NO	= oRs1("AGENT_RGST_NO")
		REQUEST_DT		= oRs1("REQUEST_DT")
		APPO_NO			= oRs1("APPO_NO")
		APPO_DT			= oRs1("APPO_DT")
		APPO_DESC		= oRs1("APPO_DESC")
				
		TAX_DOC_CD		= oRs1("TAX_DOC_CD")	' -- ����Ÿ�� ���� �����ڵ尡 �޶����Ƿ� 

		Call SubCloseRs(oRs1)	

		If EX_RECON_FLG = "Y" Then
			' -- ��1ȣ������ �о�´�.
			Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

			' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
			gCursorLocation = adUseClient 

			If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
				Exit Function
			End If
		End If

		PrintLog "LoadData Success "
		
		LoadData = True
	End Function				

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.MoveFirst
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFirst()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm)
		Else
			GetData		= ""
		End If
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
	End Sub


	' ------------------ ��ȸ �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
				lgStrSQL = lgStrSQL & " (	SELECT  "
	            lgStrSQL = lgStrSQL & "		TOP 1 TAX_DOC_CD " & vbCrLf
	            lgStrSQL = lgStrSQL & "		FROM TB_TAX_DOC " & vbCrLf	' ����ڱ��Ѻ� �޴��� 
				lgStrSQL = lgStrSQL & "		WHERE TAX_DOC_CD IN ('A100', 'A138') " & vbCrLf            
	            lgStrSQL = lgStrSQL & " ) TAX_DOC_CD " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY	A  WITH (NOLOCK) " & vbCrLf	' ����ڱ��Ѻ� �޴��� 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

	      Case "H2"
				lgStrSQL = " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_AGENT_INFO	A  WITH (NOLOCK) " & vbCrLf	' ����ڱ��Ѻ� �޴��� 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
		End Select
		PrintLog "SubMakeSQLStatements_WB101MA1 : " & lgStrSQL
	End Sub	
End Class



' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_WB101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, sNowDt, blnError, iSeqNo
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    
    blnError = False
    PrintLog "MakeHTF_WB101MA1 IS RUNNING: "
	lgsPGM_ID	= "WB101MA1"

	Set lgcCompanyInfo = New C_COMPANY_HISTORY		' -- ��� Include���� ����� ���α��� Ŭ���� 

	If Not lgcCompanyInfo.LoadData Then Exit Function			' -- ���α������� �ε� 
		
	'==========================================
	' -- ���� ���� ���� 
	'sNowDt = UNI8Date(Date())
	
	Call InitFileSystem("../../files/" & wgCO_CD ,"HomeTaxFile_" & wgCO_CD &".A100") 	
	

		
	'==========================================
	' -- ���� �Ű� ���� ���� �� �������� 
	sHTFBody = "81"
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	sHTFBody = sHTFBody & "31"

	If Not ChkNotNull(lgcCompanyInfo.FISC_END_DT, "�����������") Then blnError = True
	sHTFBody = sHTFBody & UNI6Date(lgcCompanyInfo.FISC_END_DT)

	If Not ChkBoundary("81,82,84,86,87,88", GetRgstNo42(lgcCompanyInfo.OWN_RGST_NO), "����ڵ�Ϲ�ȣ(4:2)") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO), 13)
	
	If lgsREP_TYPE = "2" Then	' -- �Ű��� 
		sHTFBody = sHTFBody & "3"
	Else
		sHTFBody = sHTFBody & "1"
	End If
	sHTFBody = sHTFBody & "0001"	' -- �Ű����� 
	sHTFBody = sHTFBody & "0001"	' -- ������ȣ 
	sHTFBody = sHTFBody & "8"	' -- �����ڱ��� 
	
	If Not ChkNotNull(lgcCompanyInfo.HOME_TAX_USR_ID, "�����ID") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_USR_ID, 20)

	If Not ChkNotNull(lgcCompanyInfo.INCOM_DT, "�Ű� ������") Then blnError = True
	sHTFBody = sHTFBody & UNI6Date(lgcCompanyInfo.INCOM_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.LAW_RGST_NO, "���ε�Ϲ�ȣ") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.LAW_RGST_NO), 13)
 
	If Not ChkNotNull(lgcCompanyInfo.CO_NM, "���θ�") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.CO_NM, "���θ�") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.CO_NM, 60)
	
	If Not ChkNotNull(lgcCompanyInfo.REPRE_NM, "��ǥ�ڸ�") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.REPRE_NM, "��ǥ�ڸ�") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.REPRE_NM, 30)
	
	If Not ChkNotNull(lgcCompanyInfo.CO_ADDR, "����������") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.CO_ADDR, "����������") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.CO_ADDR, 70)
	
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_E_MAIL, 30)	' -- �̸����ּ� 
	
	If Not ChkTelNo(lgcCompanyInfo.TEL_NO, "�������ȭ��ȣ") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TEL_NO, 14)	' -- ��ȭ��ȣ 
	
	If Not ChkNotNull(lgcCompanyInfo.IND_CLASS, "����") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.IND_CLASS, 30)
	
	If Not ChkNotNull(lgcCompanyInfo.IND_TYPE, "����") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.IND_TYPE, 50)

	If Not ChkNotNull(lgcCompanyInfo.HOME_TAX_MAIN_IND, "�־����ڵ�") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_MAIN_IND, 7)
	
	If Not ChkNotNull(lgcCompanyInfo.FISC_START_DT, "����������") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.FISC_START_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.FISC_END_DT, "�����������") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.FISC_END_DT)

	If Not ChkDateFrTo(lgcCompanyInfo.FISC_START_DT, lgcCompanyInfo.FISC_END_DT, "����������", "�����������", False) Then blnError = True
	
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_ANY_START_DT) ' -- ���úΰ��Ⱓ 
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_ANY_END_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.HOME_FILE_MAKE_DT, "�Ű��ۼ�����") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_FILE_MAKE_DT)
	
	If Not ChkContents(lgcCompanyInfo.AGENT_NM, "�����븮�μ���") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_NM, 30)	' -- �����븮�μ��� 
	
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_BAN_NO), 6)	' -- �����븮�ΰ�����ȣ: UNIRemoveDash (2006.03.07����)
	
	If Not ChkTelNo(lgcCompanyInfo.AGENT_TEL_NO, "�����븮����ȭ��ȣ") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_TEL_NO, 14)	' -- �����븮����ȭ��ȣ 
	
	sHTFBody = sHTFBody & UNIChar("1031", 4)	' -- SDS uniERP �������α׷��ڵ� 
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.AGENT_RGST_NO), 10)	' -- �����븮�μ��� 

	sHTFBody = sHTFBody & UNIChar("", 19)	' -- ���� 

	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If


	If lgcCompanyInfo.EX_RECON_FLG = "Y" Then
		Dim blnChk
		
		If Mid(lgcCompanyInfo.RECON_MGT_NO, 4, 1) = "8" Then
			blnChk = True	' -- �������� : 8 �� �ƴ� ��� �׸����߿� �����ڰ�����ȣ/�����븮�λ���ڹ�ȣ�� �����ؾ� �Ѵ�.
		Else
			blnChk = False	' -- �������� 
		End If
		
		 '-- 2006.03 �����ݽ�û�� ���� 
		sHTFBody = "83"
		sHTFBody = sHTFBody & "A218"	' 
		 
		If Not ChkNotNull(lgcCompanyInfo.GetData("W_NAME"), "��ǥ��_����") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_NAME"), 30)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO1"), "��ǥ��_��Ϲ�ȣ") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO1")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_MGT_NO"), "��ǥ��_������ȣ") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_MGT_NO")), 6)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO"), "��ǥ��_����ڵ�Ϲ�ȣ") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO2"), "��ǥ��_�ֹε�Ϲ�ȣ") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO2")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_CO_ADDR"), "��ǥ��_����������") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_CO_ADDR"), 70)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_HOME_ADDR"), "��ǥ��_�ּ�") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_HOME_ADDR"), 70)

		If Not ChkNotNull(lgcCompanyInfo.REQUEST_DT, "��û����") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.REQUEST_DT)

		If Not ChkNotNull(lgcCompanyInfo.APPO_NO, "������ȣ") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.APPO_NO), 5)

		If Not ChkContents(lgcCompanyInfo.APPO_DESC, "��������") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.APPO_DESC, 50)

		If Not ChkNotNull(lgcCompanyInfo.APPO_DT, "��������") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.APPO_DT)
		
		sHTFBody = sHTFBody & UNIChar("", 8) & vbCrLf	' -- ���� 

		If Not blnError Then
			PrintLog "WriteLine2File : " & sHTFBody
			Call PushRememberDoc(sHTFBody)	' -- �ٷ� ������� �ʰ� ����Ų��(inc_HomeTaxFunc.asp�� ����)
		End If

		If blnChk = False Then
			If lgcCompanyInfo.GetData("W_MGT_NO") = lgcCompanyInfo.RECON_MGT_NO Or lgcCompanyInfo.AGENT_RGST_NO = lgcCompanyInfo.GetData("W_RGST_NO") Then
				blnChk = True	' -- ��ġ �߰� 
			End If
		End If
		
		' -- ����������..
		lgcCompanyInfo.MoveNext 
		iSeqNo = 1	: sHTFBody = ""	' -- �ʱ�ȭ 
		
		' -- ������ �׸��� 
		Do Until lgcCompanyInfo.EOF 
			 '-- 2006.03 �����ݽ�û�� ���� 
			sHTFBody = sHTFBody & "84"
			sHTFBody = sHTFBody & "A218"	' 
			
			sHTFBody = sHTFBody & UNINumeric(iSeqNo, 6, 0)
			
			If Not ChkNotNull(lgcCompanyInfo.GetData("W_NAME"), "������_����") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_NAME"), 30)
			
			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO1"), "������_��Ϲ�ȣ") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO1")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_MGT_NO"), "������_������ȣ") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_MGT_NO")), 6)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO"), "������_����ڵ�Ϲ�ȣ") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO2"), "������_�ֹε�Ϲ�ȣ") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO2")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_CO_ADDR"), "������_����������") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_CO_ADDR"), 70)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_HOME_ADDR"), "������_�ּ�") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_HOME_ADDR"), 70)

			sHTFBody = sHTFBody & UNIChar("", 23) & vbCrLf	' -- ���� 

			' -- �����ʿ� 
			If blnChk = False Then
				If lgcCompanyInfo.GetData("W_MGT_NO") = lgcCompanyInfo.RECON_MGT_NO Or lgcCompanyInfo.AGENT_RGST_NO = lgcCompanyInfo.GetData("W_RGST_NO") Then
					blnChk = True	' -- ��ġ �߰� 
				End If
			End If
			
			iSeqNo = iSeqNo + 1
			lgcCompanyInfo.MoveNext 
		Loop

		If blnChk = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "�����븮�� �⺻������ �����븮�λ���ڹ�ȣ�� 4��° �ڸ��� '8'�� �ƴ� ��� �����ݿ� ���� ����(�׸���)�߿� ������ȣ/����� ��ȣ�� ��ġ�ϴ� ����Ÿ�� �����ؾ� �մϴ�.")
		End If
		
		If Not blnError Then
			PrintLog "WriteLine2File : " & sHTFBody
			Call PushRememberDoc(sHTFBody)	' -- �ٷ� ������� �ʰ� ����Ų��(inc_HomeTaxFunc.asp�� ����)
		End If

		' -- ���ڵ�� �ݱ� 
		Call lgcCompanyInfo.CloseRs
	End If
	
	' -- ���ΰ����� �ڿ����� ���Ƿ� �޸������� ���� 
End Function


%>
