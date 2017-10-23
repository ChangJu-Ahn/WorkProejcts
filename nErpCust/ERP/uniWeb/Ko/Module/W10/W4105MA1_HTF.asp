
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��5ȣ Ư������������� 
'*  3. Program ID           : W4105MA1
'*  4. Program Name         : W4105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_5

Set lgcTB_5 = Nothing ' -- �ʱ�ȭ 

Class C_TB_5
	' -- ���̺��� �÷����� 
	Dim SEQ_NO
	Dim W1_CD
	Dim W1
	Dim W2
	Dim W2_CD
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W7
	Dim W8
	
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
			SEQ_NO		= lgoRs1("SEQ_NO")
			W1_CD		= lgoRs1("W1_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_CD		= lgoRs1("W2_CD")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			W7			= lgoRs1("W7")
			W8			= lgoRs1("W8")
		Else
			SEQ_NO		= ""
			W1_CD		= ""
			W1			= ""
			W2			= ""
			W2_CD		= ""
			W3			= 0
			W4			= 0
			W5			= 0
			W6			= 0
			W7			= 0
			W8			= 0
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
				lgStrSQL = lgStrSQL & " FROM TB_5 A WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W4105MA1
	Dim W4101MA1
	Dim W4103MA1
	Dim A140
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W4105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim arrdblSumAmt(8), dbl19Amt, dbl17Amt, dbl18Amt, dbl40Amt, dbl46Amt,dblW19Amt,dblW40Amt
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W4105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W4105MA1"

	Set lgcTB_5 = New C_TB_5		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_5.LoadData Then Exit Function			
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W4105MA1
	
	'==========================================
	' -- ��5ȣ Ư������������� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	' -- �ּ�: ����û���ϻ�������� �������ĳ��븸 �����ϴ°� �ƴ϶� �������� ����Ÿ�ʵ嵵 �����Ѵ�.
	
	lgcTB_5.Find "W2_CD = '01'"	' -- �߼ұ�������غ��(2006.03 �������Ŀ� ����)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- ȸ����� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- �ѵ��ʰ��� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- ������ 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- �����Ѽ�����ձݺ��ξ� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- �ձݺһ��԰� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- �ձݻ��԰� 

	lgcTB_5.Find "W2_CD = '42'"	' -- �ֱǻ����߼ұ�� ���� ����ս��غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '02'"	' -- �����η°����غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '03'"	' -- �����ڼս��غ��(2006.03 �������Ŀ� ����)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- ȸ����� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- �ѵ��ʰ��� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- ������ 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- �����Ѽ�����ձݺ��ξ� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- �ձݺһ��԰� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- �ձݻ��԰� 

	lgcTB_5.Find "W2_CD = '08'"	' -- ��ȸ�����ں������غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '16'"	' -- ���������������ȸ��(2006.03 �������Ŀ� ����)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- ȸ����� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- �ѵ��ʰ��� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- ������ 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- �����Ѽ�����ձݺ��ξ� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- �ձݺһ��԰� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- �ձݻ��԰� 

	lgcTB_5.Find "W2_CD = '45'"	' -- �ε�������ȸ���� ���ڼս��غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '17'"	' -- 100%�ձݻ��԰������� ����غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '18'"	' -- �ε�������ȸ���� ���ڼս��غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '14'"	' -- ���밳�������غ��(2006.03 �������Ŀ� ����)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- ȸ����� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- �ѵ��ʰ��� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- ������ 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- �����Ѽ�����ձݺ��ξ� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- �ձݺһ��԰� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- �ձݻ��԰� 

	lgcTB_5.Find "W2_CD = '43'"	' -- �ֱǻ����� ���� �ڻ��� ó�мս��غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '47'"	' -- ��ȭ����غ�� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '19'"	' -- �غ�� �� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '40'"	' -- Ư�������󰢺� �� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '46'"	' -- Ư���ڻ갨���󰢺� �� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '41'"	' -- �غ�� �� Ư�������󰢺� �� 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '44'"	' -- ��������� 
	sHTFBody = sHTFBody & UNIChar(lgcTB_5.W1, 50)
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)


	lgcTB_5.Find "W2_CD = '17'"	' -- 100%�ձݻ��԰������� ����غ�� 
	If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)

	lgcTB_5.Find "W2_CD = '18'"	' -- 80%�ձݻ��԰������� ����غ�� 
	If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 14)	' -- ���� 2006.03.07 ���� 
	

	' ���� ������ ���� ���� ���� (���� ���� ������ �ּ�ó�����)
	' ------------------------------------------------------------------------------------
	lgcTB_5.MoveFirst
	
	Do Until lgcTB_5.EOF 
	
		If lgcTB_5.W2_CD <> "" Then
			
			Select Case UNICDbl(lgcTB_5.W2_CD , 0)
				Case 01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43
						
					' -- �����׸� 
					' (5)������ = (3)ȸ����� - (4)�ѵ��ʰ��� 
					If UNICDbl(lgcTB_5.W5, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W4, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(5)������","(3)ȸ����� - (4)�ѵ��ʰ���"))
					End If
			
					' (7)�ձݺһ��԰� = (4)�ѵ��ʰ��� + (6)�����Ѽ�����ձݺ��� 
					If UNICDbl(lgcTB_5.W7, 0) <> (UNICDbl(lgcTB_5.W4, 0) + UNICDbl(lgcTB_5.W6, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(7)�ձݺһ��԰�","(4)�ѵ��ʰ��� + (6)�����Ѽ�����ձݺ���"))
					End If
					
					' (8) �ձݻ��԰� = (3)ȸ����� - (7)�ձݺһ��԰� 
					If UNICDbl(lgcTB_5.W8, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W7, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(8)�ձݻ��԰�","(3)ȸ����� - (7)�ձݺһ��԰�"))
					End If
					
					arrdblSumAmt(3) = arrdblSumAmt(3) + UNICDbl(lgcTB_5.W3, 0)
					arrdblSumAmt(4) = arrdblSumAmt(4) + UNICDbl(lgcTB_5.W4, 0)
					arrdblSumAmt(5) = arrdblSumAmt(5) + UNICDbl(lgcTB_5.W5, 0)
					If lgcTB_5.W2_CD = "17" Or lgcTB_5.W2_CD = "18"  Then
					Else
						arrdblSumAmt(6) = arrdblSumAmt(6) + UNICDbl(lgcTB_5.W6, 0)
					End If
					arrdblSumAmt(7) = arrdblSumAmt(7) + UNICDbl(lgcTB_5.W7, 0)
					arrdblSumAmt(8) = arrdblSumAmt(8) + UNICDbl(lgcTB_5.W8, 0)
				Case 44
				   
				   
				   
				   'sHTFBody = sHTFBody & UNIChar(lgcTB_5.W1, 50)
				   'sHTFBody = sHTFBody & UNIChar(lgcTB_5.W2, 40)
				
				   ' -- �����׸� 
					' (5)������ = (3)ȸ����� - (4)�ѵ��ʰ��� 
					If UNICDbl(lgcTB_5.W5, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W4, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(5)������","(3)ȸ����� - (4)�ѵ��ʰ���"))
					End If
			
					' (7)�ձݺһ��԰� = (4)�ѵ��ʰ��� + (6)�����Ѽ�����ձݺ��� 
					If UNICDbl(lgcTB_5.W7, 0) <> (UNICDbl(lgcTB_5.W4, 0) + UNICDbl(lgcTB_5.W6, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(7)�ձݺһ��԰�","(4)�ѵ��ʰ��� + (6)�����Ѽ�����ձݺ���"))
					End If
					
					' (8) �ձݻ��԰� = (3)ȸ����� - (7)�ձݺһ��԰� 
					If UNICDbl(lgcTB_5.W8, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W7, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(8)�ձݻ��԰�","(3)ȸ����� - (7)�ձݺһ��԰�"))
					End If
					
					arrdblSumAmt(3) = arrdblSumAmt(3) + UNICDbl(lgcTB_5.W3, 0)
					arrdblSumAmt(4) = arrdblSumAmt(4) + UNICDbl(lgcTB_5.W4, 0)
					arrdblSumAmt(5) = arrdblSumAmt(5) + UNICDbl(lgcTB_5.W5, 0)
					If lgcTB_5.W2_CD = "17" Or lgcTB_5.W2_CD = "18"  Then
					Else
						arrdblSumAmt(6) = arrdblSumAmt(6) + UNICDbl(lgcTB_5.W6, 0)
					End If
					arrdblSumAmt(7) = arrdblSumAmt(7) + UNICDbl(lgcTB_5.W7, 0)
					arrdblSumAmt(8) = arrdblSumAmt(8) + UNICDbl(lgcTB_5.W8, 0)
					
					
					
				Case 19
					' -- �غ�ݰ� 
					If UNICDbl(lgcTB_5.W3, 0) <> arrdblSumAmt(3) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(3)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
					
					If UNICDbl(lgcTB_5.W4, 0) <> arrdblSumAmt(4) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(4)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
					
					If UNICDbl(lgcTB_5.W5, 0) <> arrdblSumAmt(5) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(5)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
					
					If UNICDbl(lgcTB_5.W6, 0) <> arrdblSumAmt(6) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(6)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
					
					If UNICDbl(lgcTB_5.W7, 0) <> arrdblSumAmt(7) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(7)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
					
					If UNICDbl(lgcTB_5.W8, 0) <> arrdblSumAmt(8) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�غ�� ��_(8)ȸ�����","����(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)ȸ������� �հ�"))
					End If
	
					' -- 200603 ����:  Ư�������������(A108)������ �ڵ�(19)�غ�ݰ��� �׸�(5)�����Ѽ�����ձݺ��ξ��������Ѽ������꼭(A140)�� �ڵ�(5)�غ���� �׸�(4)�������� ��ġ�ϴ��� ���� �߰� 
					Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp �� ���ǵ� 
											
					' -- �߰� ��ȸ������ �о�´�.
					cDataExists.A140.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
					cDataExists.A140.WHERE_SQL = " AND A.W1 = '05' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
					If Not cDataExists.A140.LoadData() Then
	
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� �����Ѽ������꼭(A140) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
					Else

						If UNICDbl(lgcTB_5.W6, 0)  <> UNICDbl(cDataExists.A140.GetData("W4") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W6 & " <> " & cDataExists.A140.GetData("W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��5ȣ Ư�������������(A108)������ �ڵ�(19)�غ�ݰ��� �׸�(5)�����Ѽ�����ձݺ��ξ�","��4ȣ �����Ѽ������꼭(A140)�� �ڵ�(5)�غ���� �׸�(4)������"))
						End If
													
					End If

					' -- ����� Ŭ���� �޸� ���� 
					Set cDataExists.A140 = Nothing		
    
			End Select
			
			Select Case UNICDbl(lgcTB_5.W2_CD , 0)
				Case 01

					Set cDataExists.W4101MA1 = new C_TB_31_1	' -- W4101MA1_HTF.asp �� ���ǵ� 
								
					' -- �߰� ��ȸ������ �о�´�.
					cDataExists.W4101MA1.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
					cDataExists.W4101MA1.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
					If Not cDataExists.W4101MA1.LoadData() Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� ���װ�����û��(A165) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
					Else

						If UNICDbl(lgcTB_5.W3, 0) > 0 AND UNICDbl(lgcTB_5.W3, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W5") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�߼ұ�������غ��_ȸ�����","��31ȣ(1)�߼ұ�������غ������_(5)ȸ�����"))
						End If

						If UNICDbl(lgcTB_5.W4, 0) > 0 AND UNICDbl(lgcTB_5.W4, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W6") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�߼ұ�������غ��_�ѵ��ʰ���","��31ȣ(1)�߼ұ�������غ������_(6)�ѵ��ʰ���"))
						End If

						If UNICDbl(lgcTB_5.W6, 0) > 0 AND UNICDbl(lgcTB_5.W6, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W7") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�߼ұ�������غ��_�����Ѽ�����ձݺ��ξ�","��31ȣ(1)�߼ұ�������غ������_(7)�����Ѽ����뿡�����ձݺ��ξ�"))
						End If												
					End If

					' -- ����� Ŭ���� �޸� ���� 
					Set cDataExists.W4101MA1 = Nothing		
					
					

				Case 02

					Set cDataExists.W4103MA1 = new C_TB_31_2	' -- W4103MA1_HTF.asp �� ���ǵ� 
								
					' -- �߰� ��ȸ������ �о�´�.
					cDataExists.W4103MA1.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
					cDataExists.W4103MA1.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
					If Not cDataExists.W4103MA1.LoadData() Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� ���װ�����û��(A165) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
					Else

						If UNICDbl(lgcTB_5.W3, 0) > 0 AND UNICDbl(lgcTB_5.W3, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W4") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������η°����غ��_ȸ�����","��31ȣ(2)�������η°����غ��_(4)ȸ�����"))
						End If

						If UNICDbl(lgcTB_5.W4, 0) > 0 AND UNICDbl(lgcTB_5.W4, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W5") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������η°����غ��_�ѵ��ʰ���","��31ȣ(2)�������η°����غ��_(5)�ѵ��ʰ���"))
						End If

						If UNICDbl(lgcTB_5.W6, 0) > 0 AND UNICDbl(lgcTB_5.W6, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W6") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������η°����غ��_�����Ѽ�����ձݺ��ξ�","��31ȣ(2)�������η°����غ��_(6)�����Ѽ����뿡�����ձݺ��ξ�"))
						End If												
					End If

					' -- ����� Ŭ���� �޸� ���� 
					Set cDataExists.W4103MA1 = Nothing		
				
				Case 19
					dbl19Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 17
					dbl17Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 18
					dbl18Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 40
					dbl40Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 46
					dbl46Amt = 	UNICDbl(lgcTB_5.W5, 0)
					
			End Select
			
			If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
		    
		    
		    If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '�����Ѽ����� 
					
				If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
			End IF	
					
			If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)
		End If
		
		lgcTB_5.MoveNext 
	Loop

	' -- ���� 
	If dblW19Amt - dbl17Amt - dbl18Amt > 0 Then
	
		Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp �� ���ǵ� 
								
		' -- �߰� ��ȸ������ �о�´�.
		cDataExists.A140.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A140.WHERE_SQL = " AND A.W1 = '05' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
		If Not cDataExists.A140.LoadData() Then
	
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� �����Ѽ������꼭(A140) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else

			If dblW19Amt - dbl17Amt - dbl18Amt <> UNICDbl(cDataExists.A140.GetData("W3") , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������η°����غ��_ȸ�����","��31ȣ(2)�������η°����غ��_(4)ȸ�����"))
			End If
										
		End If

		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A140 = Nothing		
	End If

	If dblW40Amt + dbl46Amt > 0 Then
	
		Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp �� ���ǵ� 
								
		' -- �߰� ��ȸ������ �о�´�.
		cDataExists.A140.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A140.WHERE_SQL = " AND A.W1 = '06' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
		If Not cDataExists.A140.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� �����Ѽ������꼭(A140) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else

			If dblW19Amt - dbl17Amt - dbl18Amt <> UNICDbl(cDataExists.A140.GetData("W3") , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������η°����غ��_ȸ�����","��31ȣ(2)�������η°����غ��_(4)ȸ�����"))
			End If
										
		End If

		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A140 = Nothing		
	End If

	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_5 = Nothing	' -- �޸����� 
	
End Function


%>
