<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��8ȣ(��) �������鼼�� ���� 
'*  3. Program ID           : W6125MA1
'*  4. Program Name         : W6125MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_8A

Set lgcTB_8A = Nothing ' -- �ʱ�ȭ 

Class C_TB_8A
	' -- ���̺��� �÷����� 
	Dim W_TYPE
	Dim W1_CD
	Dim W1
	Dim W2
	Dim W2_1
	Dim W3
	Dim W4
	Dim W7
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim SELECT_SQL		' -- ���� �ʵ� ���� 
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
		lgoRs1.MoveFirst
		lgoRs1.Find pWhereSQL
		Call GetData
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W_TYPE		= lgoRs1("W_TYPE")
			W1_CD		= lgoRs1("W1_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_1		= lgoRs1("W2_1")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W7			= lgoRs1("W7")
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

				If SELECT_SQL = "" Then
					lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & SELECT_SQL & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_8A A WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				lgStrSQL = lgStrSQL & "	ORDER BY W_TYPE , W1_CD "
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6125MA1
	Dim A161
	Dim A174
	Dim A165
	Dim A101
	Dim A179
	Dim A181
	Dim A175
	Dim A149
	Dim A154
	Dim A151
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6125MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim dblAmt, dblSum1Amt, dblSum2Amt, dblSum3Amt, dbl10Amt, dbl66Amt, dbl30Amt, dbl49Amt, dbl70Amt, dbl80Amt, dbl50Amt
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6125MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6125MA1"

	Set lgcTB_8A = New C_TB_8A		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_8A.LoadData Then Exit Function			
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W6125MA1
	
	'==========================================
	' -- ��8ȣ(��) �������鼼�� ���� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	' -- 2006.03.07 ����: ȭ������� ���ϻ��� ������ �ٸ��� 
	
	lgcTB_8A.Find "W2_1 = '01'"	' -- �ܱ����μ��װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '02'"	' -- ���ؼսǼ��װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '71'"	' -- ����ҵ漼���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '03'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '90'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '08'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '09'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '04'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '07'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '94'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '86'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '87'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = 88''"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '72'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '73'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '81'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '82'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '61'"	' -- ���װ��� 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�61 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)

	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '10'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- �׸���2 
	lgcTB_8A.Find "W2_1 = '11'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '74'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '12'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '13'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '23'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '14'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '16'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '17'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '18'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '19'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '24'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '97'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '98'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '64'"	' -- ���װ��� 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�64 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)
	
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '30'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- �׸��� 3
	lgcTB_8A.Find "W2_1 = '31'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '93'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '75'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '32'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '85'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '34'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '76'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '35'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '36'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '77'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '37'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '91'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '42'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '95'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '96'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '65'"	' -- ���װ��� 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�65 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)
	
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '49'"	' -- ���װ���2_�Ұ�_�������� 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '50'"	' -- ���װ���2_�հ�_�������� 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '51'"	' -- �������鼼���Ѱ� 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '83'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '89'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	' -- ���������� �߰��Ȱ͵�(��.��)
	' -- �׸���1�� �߰��Ȱ͵� 
	lgcTB_8A.Find "W2_1 = '57'"	' -- ������������ ���߻�������� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '58'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '59'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '60'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '67'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '66'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '99'"	' -- ���װ��� 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�99 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)

	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- �׸���3�� �߰��Ȱ� 
	lgcTB_8A.Find "W2_1 = '84'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '70'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '80'"	' -- ���װ��� 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 14)	' -- ����	2006.03.07 ���� 
	

	' -- ���� ������ �������� �����Ѵ�	
	'lgcTB_8A.MoveFirst		' -- �� �������δ� �̻��� ������ �̵��Ѵ�. 2006.03.09
	lgcTB_8A.Find "W2_1 = '03'"	' -- ���װ��� 
	' -----------------------------------------------------------------
	
	Do Until lgcTB_8A.EOF 
	
		Select Case lgcTB_8A.W_TYPE 
			Case "0"	' W3, W4 ��� 
				
				' -- ����Ÿ ���� 
				If lgcTB_8A.W2_1 = "61" Then		' -- ����ڰ� ������ �����Ҽ� �ִ� �ڵ� 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�61 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
					End If
					
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)
				End If	
				if lgcTB_8A.W2_1 <> 10 then '�Ұ� 
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
				End if
				If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
			
				' -- �� �ڵ庰 ���� 
				Select Case lgcTB_8A.W2_1
					Case "02"
						'�������鼼�װ�꼭(1)(A161)�� ���ؼսǼ��װ��� �׸�(4)���ؼսǼ��װ����� �������鼼�װ� ��ġ 
						'(�ڵ�(2)�� �׸�(4)�� ��0����Ÿ ū ��� A161 �ݵ�� �Է�)
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- ���鼼�� 
					
							Set cDataExists.A161 = new C_TB_8_1	' -- W6124MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A161.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A161.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A161.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_���������� 0 ���� ū ��� �������鼼�װ�꼭(1)(A161) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A161.GetData("W4_2") , 0) Then
									blnError = True
								
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A161 = Nothing		
					
						End If
						
					Case "90" ' -- ���â����â������� ���� ���װ���_��󼼾�(�ڵ�90)

						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- ��󼼾� 
					
							'�ڵ�(90)�� ��󼼾��� "0"���� ū ��� ���â����â��������鼼�װ�꼭(A174)�� �������� ���� 
							
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(90)�� ���â����â������� ���� ���װ��� ��󼼾��� 0���� ū ��� ���â����â��������鼼�װ�꼭(A174)�� ������ �־���մϴ�.�ش缭���� �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If

					Case "08"
					   
					   '�ڵ�(08)�� ��󼼾��� 0���� ū ��� ����׺��縦 �����ǻ�Ȱ�������� �������� �����ϴ� ���ο� ���� �ӽ�Ư�����װ����û��(A207)�� �������� ���� 
					    If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- ��󼼾� 
						
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(08)�� ��󼼾��� 0���� ū ��� ����׺��縦 �����ǻ�Ȱ�������� �������� �����ϴ� ���ο� ���� �ӽ�Ư�����װ����û��(A207)�� ������ �־���մϴ�. �ش缭���� �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If
				
						
					Case "04" '
					    '�ڵ�(04)�� ��󼼾��� 0���� ū ��� �������չ��μ��׸�����û��(A208)�� �������� ���� 
					
						
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then
					
						
						
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(04)�� ��󼼾��� 0���� ū ��� �������չ��μ��׸�����û��(A208)��  ������ �־���մϴ�.�ش缭���� �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If
					
					Case "09" ' -- �����ǻ�Ȱ���������� �������翡 ���� ���װ���_��󼼾�(�ڵ�09)
					    '�ڵ�(09)�� ��󼼾��� 0���� ū ��� ����׺��縦 �����ǻ�Ȱ�������� �������� �����ϴ� ���ο� ���� �ӽ�Ư�����װ����û��(A207)�� �������� ���� 
						
						
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(09)�� ��󼼾��� 0���� ū ��� ����׺��縦 �����ǻ�Ȱ�������� �������� �����ϴ� ���ο� ���� �ӽ�Ư�����װ����û��(A207)��  ������ �־���մϴ�.�ش� ������ �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If
					

					Case "07" ' -- �������չ��� ���� ���װ���_��󼼾�(�ڵ�07)
						'- �ڵ�(07)�� ��󼼾��� 0���� ū ��� �������չ��μ��׸�����û��(A209)�� �������� ���� 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- ��󼼾� 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(07)�� ��󼼾��� 0���� ū ��� �������չ��μ��׸�����û��(A209)��  ������ �־���մϴ�.�ش� ������ �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If
					
			

					Case "94" ' -- ���������������� ���� ���װ���_��󼼾�(�ڵ�94)
					    ' - �ڵ�(94)�� ��󼼾��� "0"���� ū ��� ����������������������׸���(A222)�� �������� ���� 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- ��󼼾� 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(94)�� ��󼼾��� 0���� ū ��� ����������������������׸���(A222)��  ������ �־���մϴ�.�ش� ������ �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						End If	

					' -- 200603 �������� �ݿ� 
					Case "70"
						dbl70Amt = dblSum1Amt	
						
						If UNICDbl(lgcTB_8A.W4, 0) <> dbl70Amt Then	' -- ���鼼�� 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & dbl70Amt, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�(�ڵ�70)_���鼼��","�ڵ�(03+90+08+09+04+07+86+87+88+57+58+59+60+67+72+73+81+82+61)_���鼼��"))
						End If	
						lgcTB_8A.W4 = 0	: dblSum1Amt = 0	' -- �������� �氪�� ���ϴ°� �ʱ�ȭ 

					' -- 200603 �������� �ݿ� 
					Case "80"

						dbl80Amt = dblSum1Amt

						If UNICDbl(lgcTB_8A.W4, 0) <> dbl80Amt Then	' -- ���鼼�� 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & dbl80Amt & ";" & dblSum1Amt, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�(�ڵ�80)_���鼼��","�ڵ�(01+02+71+66+94+99)_���鼼��"))
						End If	
					
						lgcTB_8A.W4 = 0	: dblSum1Amt = 0
					' -- 200603 �������� �ݿ� 
					Case "10" ' -- �Ұ� 
						If UNICDbl(lgcTB_8A.W4, 0) <> (dbl70Amt + dbl80Amt) Then	' -- ���鼼�� 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & (dbl70Amt + dbl80Amt), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�(�ڵ�10)_���鼼��","�ڵ�(70+80)_���鼼��"))
						End If	
						
						dbl10Amt = 	UNICDbl(lgcTB_8A.W4, 0)		' -- �ڵ�10+�ڵ�66 �ݾװ� 3ȣ �񱳴� �� �������� �����ۿ��� �Ѵ�.	
					
					Case "66"
						dbl66Amt = 	UNICDbl(lgcTB_8A.W3, 0)																
				End Select
				
				dblSum1Amt = dblSum1Amt + UNICDbl(lgcTB_8A.W4, 0)	' �Ұ�_���鼼�� 
				
			Case "1"
				' -- ����Ÿ ���� 
				If lgcTB_8A.W2_1 = "64" Then		' -- ����ڰ� ������ �����Ҽ� �ִ� �ڵ� 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�64 ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
					End If
									
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)
				End If	
				If  lgcTB_8A.W2_1 <> "30" Then
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
				End If	
				
				If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
			
				' -- �� �ڵ庰 ���� 
				Select Case lgcTB_8A.W2_1
				
					Case "16"	' -- ���������߼ұ������_��󼼾� 
					   '- �ڵ�(16)�� ��󼼾��� 0���� ū ��� �����ǰ��о����ǿ� ���� �������� �����ϴ� �߼ұ�����鼼�װ�꼭(A206)�� �������� ���� 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- ��󼼾� 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("�ڵ�(16)�� ��󼼾��� 0���� ū ��� �����ǰ��о����ǿ� ���� �������� �����ϴ� �߼ұ�����鼼�װ�꼭(A206)��  ������ �־���մϴ�.�ش� ������ �������� ������ UNIERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
						
						End If	

					

					
					Case "30"	' -- �Ұ�		
						If UNICDbl(lgcTB_8A.W4, 0) <> dblSum2Amt Then	' -- ���鼼�� 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�(�ڵ�30)_���鼼��","�ڵ�(11+74+12+13+23+14+16+17+18+19+24+97+98+64)_���鼼��"))
						End If	
						dbl30Amt = 	UNICDbl(lgcTB_8A.W4, 0)										
				End Select
				
				dblSum2Amt = dblSum2Amt + UNICDbl(lgcTB_8A.W4, 0)	' �Ұ�_���鼼�� 
				
				
			Case "2"		' W3, W4, W7 ��� 
				If lgcTB_8A.W2_1 = "65"  Then		' -- ����ڰ� ������ �����Ҽ� �ִ� �ڵ� 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Or UNICDbl(lgcTB_8A.W7, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "�ڵ�" & lgcTB_8A.W2_1 & " ����") Then blnError = True	' -- �ݾ��� 0�ƴϸ� Not Null
					End If
								
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- ����� ���� ����(Null ���)
				End If	
				
				If lgcTB_8A.W2_1 ="49" Or lgcTB_8A.W2_1 = "50" Or lgcTB_8A.W2_1 = "51" Then	
				Else
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
						
					If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
				End if
				
				If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)
				
				' -- 5 + 6 > 7
				If UNICDbl(lgcTB_8A.W3, 0) + UNICDbl(lgcTB_8A.W4, 0) < UNICDbl(lgcTB_8A.W7, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT,  lgcTB_8A.W1 &  "��������","�����̿���+���߻���"))
				End If
				
				' -- �� �ڵ庰 ���� 
				Select Case lgcTB_8A.W2_1
				
					
					Case "75" ' -- ����Ǿ����������������� ���װ���_���߻� 
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- ��󼼾� 
					
								

							Set cDataExists.A179 = new C_TB_JT2_2	' -- W6109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A179.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A179.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A179.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� ����Ǿ����������������Ѱ������װ�꼭(A179) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A179.GetData("W13") , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����Ǿ����������������� ���װ���","����Ǿ����������������Ѱ������װ�꼭(A179)�� �׸�(13)��������"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A179 = Nothing							
						End If	
					
' -- 2006.03.27 ��������	
'					Case "32" ' -- �����η°��ߺ񼼾װ�_���߻� 
'						PrintLog "dbl66Amt : " & dbl66Amt
'						' -- 200603 ����: 66�ڵ尡 ���� �ö� 
'						If UNICDbl(lgcTB_8A.W4, 0) + dbl66Amt > 0 Then	' -- ��󼼾� 
'					      '- �ڵ�(32)�� ����������"0"���� ū ��� �������η°��ߺ�߻�����(A181)�� �������� ���� 
'
'							Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp �� ���ǵ� 
'								
'							' -- �߰� ��ȸ������ �о�´�.
'							cDataExists.A181.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
'							cDataExists.A181.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
'								
'							If Not cDataExists.A181.LoadData() Then
'								blnError = True
'								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� �������η°��ߺ�߻�����(A181) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
'						
'							End If
'					
'							
'							' -- ����� Ŭ���� �޸� ���� 
'							Set cDataExists.A181 = Nothing							
'						End If	
						
						
						
					Case "37" ' 
						If UNICDbl(lgcTB_8A.W4, 0) <> 0 Then	' -- ��󼼾� 
					
						 '- �ڵ�(37)�� ���������� 0 �� �ƴѰ�� �����Ư���� ������� ���鼼���հ�ǥ (A151)�� �׸�(137)�ӽ����ڼ��װ��� �׸�(4)���鼼�� �� ��ġ 

							Set cDataExists.A151 = new C_TB_13	' -- W6111MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A151.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A151.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A151.LoadData() Then
								Response.End 
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� �����Ư���� ������� ���鼼���հ�ǥ (A151) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							ElSE
							   '���鼼���հ�ǥ (A151)�� �׸�(137)�ӽ����ڼ��װ��� �׸�(4)���鼼�� �� ��ġ 
							   cDataExists.A151.FIND 1, " W2_CD='137' "	
							   if UNICDbl(lgcTB_8A.W4, 0) <>  UNIcdbl(cDataExists.A151.GetData(1, "W4") ,0) Then
							      blnError = True
								  Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ӽ����ڼ��װ�����","���鼼���հ�ǥ (A151)�� �׸�(137)�ӽ����ڼ��װ��� �׸�(4)���鼼��"))
							   End if
						
							End If
	
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A151 = Nothing							
						End If	
					Case "91" ' -- �������Ư�����װ���_���߻� 
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- ��󼼾� 
					
						

							Set cDataExists.A175 = new C_TB_JT11_5	' -- W6113MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A175.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A175.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A175.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� �������Ư�����װ��� �������װ�꼭(A175) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A175.GetData("W10") , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������Ư�����װ���","�������Ư�����װ��� �������װ�꼭(A175)�� �׸�(15)��������"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A175 = Nothing							
						End If							
						
					' -- 200603 ����: 66�ڵ尡 ���� �ö�..	
'					Case "66"
'						dbl66Amt = 	UNICDbl(lgcTB_8A.W7, 0)		' -- �ڵ�10+�ڵ�66 �ݾװ� 3ȣ �񱳴� �� �������� �����ۿ��� �Ѵ�.	
'					 If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- ��󼼾� 
'					
'						
'
'							Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp �� ���ǵ� 
'								
'							' -- �߰� ��ȸ������ �о�´�.
'							cDataExists.A181.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
'							cDataExists.A181.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
'								
'							If Not cDataExists.A181.LoadData() Then
'								blnError = True
'								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� ����Ǿ����������������Ѱ������װ�꼭(A179) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
'						
'							End If
'					
'							
'							' -- ����� Ŭ���� �޸� ���� 
'							Set cDataExists.A179 = Nothing							
'						End If	
						
						
					
					Case "49"	' --�Ұ� 
					
					
						If UNICDbl(lgcTB_8A.W7, 0) <> UNICDbl(dblSum3Amt,0) Then	' -- ������ 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_8A.W7, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�(�ڵ�30)_���鼼��","�ڵ�(31+93+75+32+85+34+76+35+36+77+37+91+42+95+84+96+65)_���鼼��"))
						End If		

						If UNICDbl(lgcTB_8A.W7, 0) > 0 Then	' -- �������� 
						
						
						   ' ���װ�����������(3)(A149)�� �׸�(116)���������� �հ�� ��ġ(�ڵ�(49)�� ���������� 0 ���� ū ��� A149 �ݵ�� �Է�)
					
							Set cDataExists.A149 = new C_TB_8_3	' -- W6103MA1_HTF.asp �� ���ǵ� 
								
							
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A149.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
						'	cDataExists.A149.WHERE_SQL = " AND A.W_CODE = '30' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
						

							If Not cDataExists.A149.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_��󼼾��� '0'���� ū ��� ���װ�����������(3)(A149) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
								
								' -- ��ġüũ���� �ʰ� ���������� üũ��: 2006.03.10 �����ݿ� 
							'Else 
							     'cDataExists.A149.FIND 2, "SEQ_NO = 9999999 "
					
							     'if UNICDbl(lgcTB_8A.W7, 0)  <>   UNICDbl(cDataExists.A149.GetData(2, "C_W116"),0)  Then
							     '  	blnError = True
								'	Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������_�Ұ�","���װ�����������(3)(A149)�� �׸�(116)���������� �հ�"))
							     'End If
							   	
							End If
					
						End If	
					
						dbl49Amt = 	UNICDbl(lgcTB_8A.W7, 0)
						Set cDataExists.A149 = Nothing			
						
					' -- 200603 �������� ���� 		
					Case "50"	 ' -- �հ�		
						' -- �ڵ�(30)��������_�Ұ� + �ڵ�(49)��������_�Ұ� 
						If dbl30Amt + dbl49Amt <> UNICDbl(lgcTB_8A.W7, 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(50)��������_�Ұ�","�ڵ�(30)��������_�Ұ� + �ڵ�(49)��������_�Ұ�"))
						End If

						'�հ�_��������(�ڵ�50) - �ڵ�(66)������󼼾��� ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(17)�������鼼��(��)�� ����(��ġ) 
						Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp �� ���ǵ� 
												
						' -- �߰� ��ȸ������ �о�´�.
						'cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
						'cDataExists.A101.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
						If Not cDataExists.A101.LoadData() Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, "", "�Ұ�_���鼼��(�ڵ�50) - �ڵ�(66)������󼼾��� '0'���� ū ��� ��3ȣ���μ�����ǥ�ع׼���������꼭(A101) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
						Else
							If UNICDbl(lgcTB_8A.W7, 0) <> UNICDbl(cDataExists.A101.W17, 0) Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID,  UNICDbl(lgcTB_8A.W7, 0) - dbl66Amt , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�_���鼼��(�ڵ�50)","��3ȣ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(17)�������鼼��(��)"))
							End If
						End If
														
						' -- ����� Ŭ���� �޸� ���� 
						Set cDataExists.A101 = Nothing	

						dbl50Amt = UNICDbl(lgcTB_8A.W7, 0)
					' -- 200603 �������� ���� 		
					Case "51"	 ' -- �հ�		

						If UNICDbl(lgcTB_8A.W7, 0) <> (dbl10Amt + dbl50Amt) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID,  UNICDbl(lgcTB_8A.W7, 0) - dbl66Amt , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������鼼���Ѱ�(51)","�ڵ�(10+50)�� ��������"))
						End If
						
					Case "83" ' -- ������Դ밡�� ������������_�������� 
						If UNICDbl(lgcTB_8A.W7, 0) > 0 Then	' --������ 

							' �̰��� : ���Ĺ̰���(������Դ밡������������������(A154))
						End If																									
				End Select
				
				dblSum3Amt = dblSum3Amt + UNICDbl(lgcTB_8A.W7, 0)	' �Ұ�_���鼼�� 
		End Select
		
		lgcTB_8A.MoveNext 
	Loop

	' -- ������ �����κ� 
	If dbl10Amt > 0 Then		' -- 200603 �������� ���� 
	
		'�Ұ�_���鼼��(�ڵ�10) + �ڵ�(66)������󼼾��� ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(19)�������鼼��(��)�� ����(��ġ)
		Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp �� ���ǵ� 
								
		' -- �߰� ��ȸ������ �о�´�.
		'cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		'cDataExists.A101.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
		If Not cDataExists.A101.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "�Ұ�_���鼼��(�ڵ�10) '0'���� ū ��� ��3ȣ���μ�����ǥ�ع׼���������꼭(A101) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
			If dbl10Amt <> UNICDbl(cDataExists.A101.W19 , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ұ�_���鼼��(�ڵ�10)","��3ȣ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(19)�������鼼��(��)"))
			End If
		End If
										
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A101 = Nothing	
	End If
								
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_8A = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6125MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A166" '-- �ܺ� ���� SQL
			lgStrSQL = ""

	End Select
	PrintLog "SubMakeSQLStatements_W7109MA1 : " & lgStrSQL
End Sub
%>
