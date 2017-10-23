<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��3ȣ��2(1)(2)ǥ�ؼ��Ͱ�꼭 
'*  3. Program ID           : W1105MA1
'*  4. Program Name         : W1105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_3_3

Set lgcTB_3_3 = Nothing ' -- �ʱ�ȭ 

Class C_TB_3_3
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	
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
		Call GetData()
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
		Call GetData()
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
		Call GetData()
	End Function	
	
	Function Clone(Byref pRs)
	  Set pRs = lgoRs1.clone
	End Function
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
		Else
			W1			= ""
			W2			= ""
			W3			= ""
			W4			= ""
			W5			= 0
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
				lgStrSQL = lgStrSQL & " FROM TB_3_3	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W1105MA1
	Dim A117
	Dim A102
	Dim A100
	Dim A129
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W1105MA1()
    Dim iKey1, iKey2, iKey3,dblAmt1
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1105MA1"

	Set lgcTB_3_3 = New C_TB_3_3		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_3_3.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	Set cDataExists = new TYPE_DATA_EXIST_W1105MA1
	Call lgcTB_3_3.Clone(oRs2)
	'==========================================
	' -- ��3ȣ��2(1)(2)ǥ�ؼ��Ͱ�꼭 ���ڽŰ� �� �������� 
	sHTFBody = "83"

	If lgcTB_3_3.W1 = "1" Then
		sHTFBody = sHTFBody & UNIChar("A115", 4) ' -- �Ϲݹ��ο� 
	Else
		sHTFBody = sHTFBody & UNIChar("A116", 4) ' -- �������ο� 
	End If

	lgcTB_3_3.Find "W4 = '01'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '02'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '03'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '04'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '05'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '06'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '07'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '08'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '09'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '10'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '11'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '12'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '13'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '91'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '14'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '15'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '16'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '17'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '18'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '19'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '20'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '21'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '22'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '23'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '24'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '25'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '26'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '27'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '28'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '29'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '30'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '31'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '32'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '33'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '34'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '35'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '36'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '37'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '38'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '39'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '40'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '41'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '42'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '43'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '44'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '45'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '46'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '47'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '48'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '49'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '50'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '51'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '52'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '53'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '54'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '55'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '56'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '57'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '58'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '59'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '60'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '61'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '62'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '63'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '64'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '65'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '66'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '67'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '68'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '69'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '70'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '71'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '72'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '73'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '74'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '75'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '76'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '77'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '78'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '79'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '80'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '81'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '82'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '83'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '84'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '85'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '201'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '202'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '203'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '204'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '86'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '211'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '212'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '213'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '214'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '87'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '221'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '222'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '223'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '224'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	sHTFBody = sHTFBody & UNIChar("", 44)	' -- ���� 
	
	' -- ���������� ���ϻ��������� �и��Ѵ�. (������ ������ ���� ���� ������ Ʋ����)
	' -------------------------------------------------------------------------------	
	lgcTB_3_3.Find "W4 = '01'"	

	If lgcTB_3_3.W1 = "1" Then
		'sHTFBody = sHTFBody & UNIChar("A115", 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

		Do Until lgcTB_3_3.EOF 
			
		  SELECt Case  lgcTB_3_3.W4 
		     Case "01"
		            oRs2.MoveFirst
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '03'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
					
					'- �ڵ�(01)�����= �ڵ� 02 + 03 + 04 + 05 + 06 + 07 + 08
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(01)�����","�ڵ� 02 + 03 + 04 + 05 + 06 + 07 + 08"))
						   blnError = True	
					End If
			  Case "09"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '10'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
					
					'-  �ڵ�(09) �������= �ڵ� 10 + 14
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(09)�������","�ڵ� 10 + 14"))
						   blnError = True	
					End If	
			  
			  Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '91'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'- �ڵ�(10)��ǰ�������= �ڵ� 11 + 12 - 13 - 91
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(10)��ǰ�������","�ڵ� 11 + 12 - 13 - 91"))
						   blnError = True	
					End If	
			
			  Case "14"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '15'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '16'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  �ڵ�(14)����,����,�Ӵ�,�о�,���,��Ÿ����= �ڵ� 15 + 16 - 17 - 18
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(14)����,����,�Ӵ�,�о�,���,��Ÿ����","�ڵ� 15 + 16 - 17 - 18"))
						   blnError = True	
					End If			
					
				Case "16"
			      
					
					'- �ڵ�(16)����ѿ����Ϲݹ����� ��� �μӸ����� �������� �հ� ��ġ16 = 
					'������������(A117)�� 
					'�ڵ�(34)(�����ǰ��������)           + �����������(A118)�� �ڵ�(32)(���������)         
					'  + �Ӵ��������(A119)�� �ڵ�(17)(�Ӵ����)           + �о��������(A120)�� �ڵ�(34)(���ϼ����õ�����)       
					'    + ��ۿ�������(A121)�� �ڵ�(30)(����ѿ�ۿ���)           + ��Ÿ��������(A123)�� �ڵ�(32)(����ѿ��� 
					
					
			        	Set cDataExists.A117 = new C_TB_3_3_3	' -- W1107MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A117.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
									
							If Not cDataExists.A117.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "�μӸ���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
						
							Else
								  cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '3' AND  W4 = '34' "	 ' �����ǰ�������� 
							     if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if
							      cDataExists.A117.Filter  ""
							      cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '4' AND  W4 = '32' "	 ' ��������� 
							      if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
							      cDataExists.A117.Filter  ""
							      cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '5' AND  W4 = '17' "	 ' �Ӵ���� 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								  cDataExists.A117.Filter "W1 = '6' AND  W4 = '34' "	 ' ���ϼ����õ����� 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								  cDataExists.A117.Filter " W1 = '7' AND  W4 = '30' "	 ' ����ѿ�ۿ��� 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								   cDataExists.A117.Filter "W1 = '8' AND  W4 = '32' "	 ' ����ѿ�ۿ��� 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
					
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A117 = Nothing				
	
							If UNICDbl(lgcTB_3_3.W5, 0) <> UNICDbl(sTmp,0) Then
								   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(16)����ѿ���" ,"������������(A117)���ڵ�(34)(�����ǰ��������) " &_
								         "  + �����������(A118)�� �ڵ�(32)(���������)  + �Ӵ��������(A119)�� �ڵ�(17)(�Ӵ����)   " &_
								         "  + �о��������(A120)�� �ڵ�(34)(���ϼ����õ�����) " &_
								         "  + ��ۿ�������(A121)�� �ڵ�(30)(����ѿ�ۿ���)           + ��Ÿ��������(A123)�� �ڵ�(32)(����ѿ���"))
								   blnError = True	
							End If				
					
					
					
			    Case "19"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
				
					
					'-  �ڵ�(19)����������= �ڵ� 01 - 09	
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(19)����������","�ڵ� 01 - 09"))
						   blnError = True	
					End If		
					
				Case "20"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '21'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '29'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '30'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '31'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '32'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '33'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '34'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '35'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
				
					
					'- �ڵ�(20)�Ǹź�Ͱ�����= �ڵ� 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29 + 30 + 31 + 32 + 33+ 34 + 35
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(20)�Ǹź�Ͱ�����","�ڵ�21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29 + 30 + 31 + 32 + 33+ 34 + 35"))
						   blnError = True	
					End If	

				' -- 200603	: ���� �߰� 
				Case "35"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '201'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '202'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '203'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '204'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- �ڵ�(35)��Ÿ�Ǹź�Ͱ�����= �ڵ� 201+202+203+204
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(35) ��Ÿ�Ǹź�Ͱ�����","�ڵ� 201 + 202 + 203 + 204"))
						   blnError = True	
					End If	
					
			   Case "36"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '19'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
				
					
					'- - �ڵ�(36)��������= �ڵ� 19 - 20			
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(36)��������","�ڵ� 19 - 20	"))
						   blnError = True	
					End If	
					
			  
					
					
						
			    Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '41'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '42'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '43'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '45'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '51'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- �ڵ�(37)�����ܼ���= �ڵ� 38 + 39 + 40 + 41 + 42 + 43 + 44 + 45 + 46 + 47 + 48 + 49 + 50+ 51 + 52			
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(37)�����ܼ���","�ڵ� 38 + 39 + 40 + 41 + 42 + 43 + 44 + 45 + 46 + 47 + 48 + 49 + 50+ 51 + 52		"))
						   blnError = True	
					End If		
					

				' -- 200603	: ���� �߰� 
				Case "52"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '211'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '212'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '213'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '214'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- �ڵ�(52) ��Ÿ�����ܼ���= �ڵ� 211+212+213+214
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(52) ��Ÿ�����ܼ���","�ڵ� 211 + 212 + 213 + 214"))
						   blnError = True	
					End If	
								
		      Case "53"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '54'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '68'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '70'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- �ڵ�(53)�����ܺ��= �ڵ� 54 + 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62 + 63 + 64 + 65 + 66+ 67 + 68 + 69 + 70	
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(53)�����ܺ��","�ڵ� 54 + 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62 + 63 + 64 + 65 + 66+ 67 + 68 + 69 + 70"))
						   blnError = True	
					End If	

				' -- 200603	: ���� �߰� 
				Case "70"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '221'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '222'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '223'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '224'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- �ڵ�(52) ��Ÿ�����ܺ��= �ڵ� 211+212+213+214
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(70) ��Ÿ�����ܺ��","�ڵ� 221 + 222 + 223 + 224"))
						   blnError = True	
					End If	
					
				Case "71"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '36'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '37'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  �ڵ�(71)�������= �ڵ� 36 + 37 - 53
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(71)�������","�ڵ� 36 + 37 - 53"))
						   blnError = True	
					End If		
					
				Case "72"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '73'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '74'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '75'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '76'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  �ڵ�(72)Ư������= �ڵ� 73 + 74 + 75 + 76
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(72)Ư������","�ڵ� 73 + 74 + 75 + 76"))
						   blnError = True	
					End If			
			  Case "77"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '78'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '79'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- �ڵ�(77)Ư���ս�= �ڵ� 78 + 79
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(77)Ư���ս�","�ڵ� 19 - 20	"))
						   blnError = True	
					End If			
				Case "80"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '71'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '72'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '77'"	
					dblAmt1 = dblAmt1- UNICDbl(oRs2("W5"), 0)
				
					
					'- �ڵ�(80)���μ����������������= �ڵ� 71 + 72 - 77
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(80)���μ����������������","�ڵ� 71 + 72 - 77	"))
						   blnError = True	
					End If		
					
				Case "81"
			        	Set cDataExists.A102 = new C_TB_15	' -- W1107MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A102.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
									
							' -- ��15ȣ ���񺰼ҵ�ݾ���������(A102)���� 
							Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							Call SubMakeSQLStatements_W1105MA1("A102",iKey1, iKey2, iKey3)   
								
							cDataExists.A102.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A102.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A102.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "��15ȣ ���񺰼ҵ�ݾ���������_�ձݻ��Թ��ͱݺһ���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
								'�ڵ�(81)���μ���� : �ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �հ�ݾ� ���� ũ�� ����(�������Ͱ��������� �ƴѰ�츸 ����)
								If UNICDbl(lgcTB_3_3.W5, 0) > UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_HIGH_AMT, "�ڵ�(81)���μ����","�ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �հ�ݾ�"))
								End If
							End If
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A102 = Nothing
				Case "65"
				   if UNICDbl(lgcTB_3_3.W5, 0) >= 5000000 Then
				       	'-�ڵ�(65)��α� > 500���� �� ����αݸ���(A129) �հ�(999999) �ݾ��� <=0 ����(��,������������ �񿵸�����(50,60,70)�ΰ�� ��������)
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
						
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "��22ȣ ��αݸ���(A129)", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
							
								  Call cDataExists.A129.Find ("2","W9_CD = '99'")
								If UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"), 0)<= 0Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg("�ڵ�(65)��α� > 500���� �� ����αݸ���(A129) �հ�(999999) �ݾ��� <=0 ���� �ȵ˴ϴ�", "",""))
								End If
							End If
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing
					End if						
					
					
				' -- 2006.03.17 ����: ��.��Ÿ�� �ݾ��� �������, ��.��.�� �ݾ��� 0�̸� ����.
				Case "204", "214", "224"
					If UNICDbl(lgcTB_3_3.W5, 0) <> 0 Then

		 				oRs2.MoveFirst
						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 3) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "�ڵ�(" & lgcTB_3_3.W4 & ") ��. ��Ÿ�� �ݾ��� 0�� �ƴҰ��, �ڵ�(" & oRs2("W4") & ") " & oRs2("W3") & " �ݾ��� 0�̸� �����Դϴ�.")
						End If

						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 2) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "�ڵ�(" & lgcTB_3_3.W4 & ") ��. ��Ÿ�� �ݾ��� 0�� �ƴҰ��, �ڵ�(" & oRs2("W4") & ") " & oRs2("W3") & " �ݾ��� 0�̸� �����Դϴ�.")
						End If

						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 1) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "�ڵ�(" & lgcTB_3_3.W4 & ") ��. ��Ÿ�� �ݾ��� 0�� �ƴҰ��, �ڵ�(" & oRs2("W4") & ") " & oRs2("W3") & " �ݾ��� 0�̸� �����Դϴ�.")
						End If
					End If
			End Select 			
			
			
			
			'If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)
		
			lgcTB_3_3.MoveNext 
	   Loop
	
	Else
	
		' -- ����/����/���Ǿ� ���ο� : uniERP�� �Ⱦ��� �����̶� ���������� ���µ��� 
	
		'sHTFBody = sHTFBody & UNIChar("A116", 4)	
		
		'Do Until lgcTB_3_3.EOF 
	
			'If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)
		
			'lgcTB_3_3.MoveNext 
	   'Loop
	End If
	
	
	
	
	' ----------- 
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3_3 = Nothing	' -- �޸����� 
	
End Function


' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W1105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	   Case "A102" '-- �ܺ� ���� SQL
		
			' -- �ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �ݾ�(2)�� �հ�� ��ġ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '1'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements_W1105MA1 : " & lgStrSQL
End Sub
%>
