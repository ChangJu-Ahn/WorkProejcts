<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��10ȣ ��õ���μ��׸��� 
'*  3. Program ID           : W7101MA1
'*  4. Program Name         : W7101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_10A

Set lgcTB_10A = Nothing ' -- �ʱ�ȭ 

Class C_TB_10A
	' -- ���̺��� �÷����� 
	Dim SEQ_NO
	Dim W1
	Dim W2_1
	Dim W2_2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		On Error Resume Next                                                             '��: Protect system from crashing
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

		' ��Ƽ�������� ù���� ���� 
		Call GetData
		
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
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			SEQ_NO		= lgoRs1("SEQ_NO")
			W1			= lgoRs1("W1")
			W2_1		= lgoRs1("W2_1")
			W2_2		= lgoRs1("W2_2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
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
				lgStrSQL = lgStrSQL & " FROM TB_10A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W7101MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W7101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7101MA1"

	Set lgcTB_10A = New C_TB_10A		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_10A.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	iSeqNo = 1
	'==========================================
	' -- ��10ȣ ��õ���μ��׸��� �������� 

	Do Until lgcTB_10A.EOF() 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
		If UNICDbl(lgcTB_10A.SEQ_NO, 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_10A.W1 , "����") Then blnError = True	' �հ��� �ܿ� �ʼ�üũ 
			If Not ChkNotNull(UNIRemoveDash(lgcTB_10A.W2_1) , "�����(�ֹ�)��Ϲ�ȣ") Then blnError = True
			If Not ChkNotNull(lgcTB_10A.W2_2 , "��ȣ(����)") Then blnError = True
			If Not ChkNotNull(lgcTB_10A.W3 , "��õ¡����") Then blnError = True
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_10A.SEQ_NO, 6)
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_10A.W1, 50)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_10A.W2_1), 13)
		sHTFBody = sHTFBody & UNIChar(lgcTB_10A.W2_2, 60)
		sHTFBody = sHTFBody & UNI8Date(lgcTB_10A.W3)
		
		If Not ChkNotNull(lgcTB_10A.W4, "���ڱݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_10A.W4, 15, 0)
		If UNICDbl(lgcTB_10A.SEQ_NO, 0) <> 999999 Then		
		   If Not ChkNotNull(lgcTB_10A.W5, "����") Then blnError = True	
		End if   
		sHTFBody = sHTFBody & UNINumeric(Replace(lgcTB_10A.W5,"%",""), 5, 2)
				
		If Not ChkNotNull(lgcTB_10A.W6, "���μ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_10A.W6, 15, 0)

				
		sHTFBody = sHTFBody & UNIChar("", 22)	 & vbCrLf	' -- ���� 
		iSeqNo = iSeqNo + 1
		Call lgcTB_10A.MoveNext()	' -- 1�� ���ڵ�� 
	Loop
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)	' �������°��� WirteLine�� �ƴ� 
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_10A = Nothing	' -- �޸����� 
	
End Function


%>
