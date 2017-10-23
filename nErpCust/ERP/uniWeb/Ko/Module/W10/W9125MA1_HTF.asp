
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : �ؿ��������� ���� 
'*  3. Program ID           : W9125MA1
'*  4. Program Name         : W9125MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : LEEWOLSAN
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_A125
Dim lgcTB_A125_1
Dim lgcTB_A125_2


Set lgcTB_A125 = Nothing ' -- �ʱ�ȭ 
Set lgcTB_A125_1 = Nothing ' -- �ʱ�ȭ 
Set lgcTB_A125_2 = Nothing ' -- �ʱ�ȭ 
Set lgcCompanyInfo = Nothing ' -- �ʱ�ȭ 




'===========================================================================
'C_TB_A125
'===========================================================================
  
  
Class C_TB_A125
	' -- ���̺��� �÷����� 
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm).value
		End If
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A125	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W6 <> '' "  & vbCrLf	' -- ����Ÿ�� ���� ���� 
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9125MA1
	Dim A126
End Class


'===========================================================================
'C_TB_A125_1
'===========================================================================
  
Class C_TB_A125_1
	' -- ���̺��� �÷����� 
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, "_������Ȳ", TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm).value
		End If
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A125_1	A  WITH (NOLOCK) JOIN TB_A125 B  " & vbCrLf	' 
				lgStrSQL = lgStrSQL & " ON A.CO_CD=B.CO_CD AND A.FISC_YEAR= B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO = B.SEQ_NO" 	 & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W2 + A.W3 + A.W4+A.W5+A.W6 <>0 " 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		ORDER BY A.SEQ_NO,SEQ " 	 & vbCrLf



				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class


'===========================================================================
'C_TB_A125_2
'===========================================================================
  
Class C_TB_A125_2
	' -- ���̺��� �÷����� 
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				'Call SaveHTFError(lgsPGM_ID, "_��ȸ����Ȳ", TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm).value
		End If
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A125_2	A  WITH (NOLOCK) JOIN TB_A125 B  " & vbCrLf	' 
				lgStrSQL = lgStrSQL & " ON A.CO_CD=B.CO_CD AND A.FISC_YEAR= B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO = B.SEQ_NO" 	 & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W4+A.W5+A.W6 <>0 " 	 & vbCrLf
				'lgStrSQL = lgStrSQL & "		AND ISNULL(W1,'')<>''" 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		ORDER BY A.SEQ_NO,SEQ " 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class


' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9125MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), blnChkA126A127
    Dim oRs3
    
'    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    PrintLog "MakeHTF_W9125MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9125MA1"

	Set lgcTB_A125 = New C_TB_A125		' -- �ش缭�� Ŭ���� 
	Set lgcTB_A125_1 = New C_TB_A125_1		' -- �ش缭�� Ŭ���� 
	Set lgcTB_A125_2 = New C_TB_A125_2		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_A125.LoadData Then Exit Function			

	Call lgcTB_A125_1.LoadData
	Call lgcTB_A125_2.LoadData

	'If Not lgcCompanyInfo.LoadData Then Exit Function			' -- ���α������� �ε� 
	
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9125MA1

	' -- �������� 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

	'==========================================

	iSeqNo = 1	: sHTFBody = ""
	
	'==========================================
	'�ؿ��������� - �⺻���� 
	'==========================================
	dim tmpVal
	Do Until lgcTB_A125.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
        '3. ��������ȣ 
        If Not ChkNotNull(lgcTB_A125.GetData("SEQ_NO"), "��������ȣ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("SEQ_NO"), 4, 0)
		
        '4 ���⸻�������μ� 
        If Not ChkNumeric(CStr(lgcTB_A125.GetData("W15")), "���⸻�������μ�") Then blnError = True
        sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W15"), 5, 0)
        
        '5 ���ż����μ� 
         If Not ChkNumeric(CStr(lgcTB_A125.GetData("W16")), "���ż����μ�") Then blnError = True
        sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W16"), 5, 0)
   
        '6 ���û����μ� 
         If Not ChkNumeric(CStr(lgcTB_A125.GetData("W17")), "���ż����μ�") Then blnError = True
         sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W17"), 5, 0)

        '7 ��⸻�������μ� 
        tmpVal= CDbl(lgcTB_A125.GetData("W15")) + cdbl(lgcTB_A125.GetData("W16")) - cDbl(lgcTB_A125.GetData("W17"))
        
         If Not ChkNumeric(tmpVal, "��⸻�������μ�") Then blnError = True

         sHTFBody = sHTFBody & UNINumeric(tmpVal, 5, 0)

        '8 �繫��Ȳǥ������μ� 
        If Not ChkNumeric(CStr(lgcTB_A125.GetData("W18")), "�繫��Ȳǥ������μ�") Then blnError = True
         sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W17"), 5, 0)

        '9 ���ڱ��ڵ� 
        If Not ChkNotNull(lgcTB_A125.GetData("W6"), "���ڱ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W6"), 3)
		
        '10 �������θ� 
        If Not ChkNotNull(lgcTB_A125.GetData("W7"), "�������θ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W7"), 60)
		
        '11 �������ΰ�����ȣ 
        If Not ChkNotNull(lgcTB_A125.GetData("W8"), "�������ΰ�����ȣ") Then blnError = True
		If Len(lgcTB_A125.GetData("W8")) <> 8 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W8"), UNIGetMesg("��ü���̰� 8�� �ƴϸ� �����Դϴ�.", "",""))
		End If
			' -- 2006.03.29 ����  = 8 ���� 
			' -- ù���ڰ� 1,2 �� �ƴϸ� 
		if lgcTB_A125.GetData("W8") <> "99999999" And (Left(lgcTB_A125.GetData("W8"), 1) <> "1" And Left(lgcTB_A125.GetData("W8"), 1) <> "2" And Left(lgcTB_A125.GetData("W8"), 1) <> "8" ) Then
		
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W8"), UNIGetMesg("�������ΰ�����ȣ�� 99999999�� �ƴҶ�, ù���ڰ� 1 �Ǵ� 2 �Ǵ� 8 ��(��) �ƴϸ� �����Դϴ�", "",""))
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W8"), 8)
		
		
        '12 �������μ����� 
        
        If Not ChkNotNull(Replace(lgcTB_A125.GetData("W9"),vbCrLf,""), "�������μ�����") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(Replace(lgcTB_A125.GetData("W9"),vbCrLf,""), 70)
		
        '13 �������� 
        If Not ChkNotNull(lgcTB_A125.GetData("W10"), "��������") Then blnError = True
		If DateDiff("m", lgcTB_A125.GetData("W10"), lgcTB_A125.GetData("W11_1")) < 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W10"), UNIGetMesg("�������ڴ� �������_�����Ϻ��� ���ų� �۾ƾ��մϴ�.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W10"))
		
       ' 14 �������_������ 
       If Not ChkNotNull(lgcTB_A125.GetData("W11_1"), "�������_������") Then blnError = True
		If DateDiff("m", lgcTB_A125.GetData("W11_1"), lgcTB_A125.GetData("W11_2")) <= 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W11_1"), UNIGetMesg("�������_�������� �������_�����Ϻ��� �۾ƾ��մϴ�.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W11_1"))
		
		
        '15 �������_������ 
        If Not ChkNotNull(lgcTB_A125.GetData("W11_1"), "�������_������") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W11_2"))
		
        '16 �����ڵ� 
        Call SubMakeSQLStatements_W9125MA1("2",lgcTB_A125.GetData("W12"), "", "")  '����üũ 
        If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("�����ڵ尡 �������� �ʽ��ϴ�.", "",""))
		End If

        
        If Not ChkNotNull(lgcTB_A125.GetData("W12"), "�����ڵ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W12"), 7)
		
		
        '17 ������ 
        If Not ChkNotNull(lgcTB_A125.GetData("W13"), "������") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A125.GetData("W13")), "������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W13"), 5, 0)
		
        '18 ������İ������� 
        
		If Not ChkNotNull(lgcTB_A125.GetData("W14"), "������İ�������") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A125.GetData("W14")), "������İ�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W14"), 5, 0)
		
        '19 û��-û���� 
        sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W19"))	' -- Null ��� 
        
        '20û��- ȸ���ݾ� 
        If CDbl(lgcTB_A125.GetData("W20")) > 0 And lgcTB_A125.GetData("W19") = "" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W16"), UNIGetMesg("ȸ���ݾ��� 0���� ũ�� û������ ����Ǿ� �־�� �մϴ�.", "",""))
		End If
		
        If IsDate(lgcTB_A125.GetData("W19")) And CDbl(lgcTB_A125.GetData("W20")) <= 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W20"), UNIGetMesg("û������ �Է½� ȸ���ݾ��� 0���� Ŀ�� �˴ϴ�.", "",""))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W20"), 15, 0)
		
        '21 ���� 60
		sHTFBody = sHTFBody & UNIChar("", 60)  & vbCrLf ' -- ����	 :

		lgcTB_A125.MoveNext 
	Loop


	'==========================================
	'2 �ؿ��������� - ������Ȳ 
	'==========================================
	
	
	Do Until lgcTB_A125_1.EOF 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
        
        '3. ��������ȣ 
         If Not ChkNotNull(lgcTB_A125_1.GetData("SEQ_NO"), "��������ȣ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("SEQ_NO"), 4, 0)
          
	
        '4.�Ϸù�ȣ 
         If Not ChkNotNull(lgcTB_A125_1.GetData("SEQ"), "�Ϸù�ȣ") Then blnError = True
        
			'--��μ�������üũ 
			IF lgcTB_A125_1.GetData("SEQ")<>"1" then

				if lgcTB_A125_1.GetData("W5")<>"0" or lgcTB_A125_1.GetData("W6")<>"0" then

					Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W5") & ":" & lgcTB_A125_1.GetData("w6"), UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "�뿩��,��μ�������",""))
					blnError = True
		      	end if
			end if
			
			IF lgcTB_A125_1.GetData("SEQ")="999999" then
				'Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W1"), UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "�뿩��,��μ�������",""))
				'blnError = True
			end if
			
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("SEQ"), 6, 0)
        
        '5.(��������)���ָ� NULL ��� 

        IF lgcTB_A125_1.GetData("SEQ")="1" then '000001�ΰ�� A100���θ�� ��ġ�ؾ���. '
       
			if lgcCompanyInfo.CO_NM<>lgcTB_A125_1.GetData("W1") then
				Call SaveHTFError(lgsPGM_ID, lgcCompanyInfo.CO_NM & ":" & lgcTB_A125_1.GetData("w1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���θ�","(��������)���ָ�"))
			end if
        ELSE
        END IF 
         'If Not ChkNotNull(lgcTB_A125_1.GetData("W1"), "(��������)���ָ�") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_1.GetData("W1"), 30)
		
        '6.���ڱݾ� 
        
        If Not ChkNotNull(lgcTB_A125_1.GetData("W2"), "���ڱݾ�") Then blnError = True	
		  sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W2"), 15, 0)
		 
        '7���ں��� 
        
			IF lgcTB_A125_1.GetData("SEQ")>"1" and lgcTB_A125_1.GetData("SEQ") < "999998"   then
				IF lgcTB_A125_1.GetData("W3")<"10" then
				Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W3"), UNIGetMesg(TYPE_CHK_OVER_EQUAL, "���ں���","10%"))
				end if
			end if
        
         If Not ChkNotNull(lgcTB_A125_1.GetData("W3"), "���ں���") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W3"), 5, 1)
		 
        '8 ���ݼ��� 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W3"), "���ݼ���") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W3"), 15, 0)
		 
        '9 �뿩�� 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W5"), "�뿩��") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W5"), 15, 0)
 
        '10 ��μ������� 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W6"), "��μ�������") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W6"), 15, 0)
		
        '11���� 39
         sHTFBody = sHTFBody & UNIChar("", 39) & vbCrLf' -- ����	 :
		
		lgcTB_A125_1.MoveNext 
		
	Loop
	

	
	'==========================================
	'3 �ؿ��������� - ��ȸ����Ȳ 
	'==========================================
	
	Do Until lgcTB_A125_2.EOF 
	
		sHTFBody =  sHTFBody & "85"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
        '3. ��������ȣ 
        If Not ChkNotNull(lgcTB_A125_2.GetData("SEQ_NO"), "��������ȣ") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("SEQ_NO"), 4, 0)

        '4.�Ϸù�ȣ 
         If Not ChkNotNull(lgcTB_A125_2.GetData("SEQ"), "�Ϸù�ȣ") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("SEQ"), 6, 0)
		
        '5.��ȸ��� 
         If Not ChkNotNull(lgcTB_A125_2.GetData("W1"), "��ȸ���") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W1"), 50)
		
        '6.�����ڵ� 
        
			Call SubMakeSQLStatements_W9125MA1("2",lgcTB_A125_2.GetData("W2"), "", "")  '����üũ 
			If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_A125_2.GetData("W2"), UNIGetMesg("�����ڵ尡 �������� �ʽ��ϴ�.", "",""))
			End If
		
		
         If Not ChkNotNull(lgcTB_A125_2.GetData("W2"), "�����ڵ�") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W2"), 7)
		
        '7 ������ 
          If Not ChkNotNull(lgcTB_A125_2.GetData("W3"), "������") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W3"), 70)


        '8 �������������ڱݾ� 
         If Not ChkNotNull(lgcTB_A125_2.GetData("W4"), "�������������ڱݾ�") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W4"), 15, 0)
		 
        '9 ���ں��� 
        If Not ChkNotNull(lgcTB_A125_2.GetData("W5"), "���ں���") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W5"), 5, 1)
		 
        '10 �������� 
        If Not ChkNotNull(lgcTB_A125_2.GetData("W6"), "��������") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W6"), 15, 0)
		 
        '11 ���� 22
         sHTFBody = sHTFBody & UNIChar("", 22)  & vbCrLf' -- ����	 :
     
		lgcTB_A125_2.MoveNext 
	
	Loop
	
	'sHTFBody = mid(sHTFBody, 1,len(sHTFBody)- 1)

	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
		
	End If
	Call SubCloseRs(oRs3)

	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_A125 = Nothing	' -- �޸����� 
	Set lgcTB_A125_1 = Nothing	' -- �޸����� 
	Set lgcTB_A125_2 = Nothing	' -- �޸����� 
	
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9125MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
		Case "A126" '-- �ܺ� ���� SQL

			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT TOP 1 1 FROM TB_A126 A "  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W01 > 0 "	 & vbCrLf	' -- �ڻ��Ѱ� 
			

		Case "1"
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT TOP 1 1 FROM TB_A125 A"  & vbCrLf
'			lgStrSQL = lgStrSQL & "		INNER JOIN TB_A125 B ON A."  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W01 > 0 "	 & vbCrLf	' -- �ڻ��Ѱ� 
		
		
		Case "2" '����üũ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " select top 1 STD_INCM_RT_CD from tb_std_income_rate"  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE ATTRIBUTE_YEAR = '2005' " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND STD_INCM_RT_CD= " & filterVar(pCode1 ,"''","S")	 & vbCrLf
	
	
	End Select
	PrintLog "SubMakeSQLStatements_W9125MA1 : " & lgStrSQL
End Sub
%>
