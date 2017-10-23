<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��8ȣ �������鼼�װ�꼭(3)
'*  3. Program ID           : W6121MA1
'*  4. Program Name         : W6121MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_8_3

Set lgcTB_8_3 = Nothing ' -- �ʱ�ȭ 

Class C_TB_8_3
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	Dim SELECT_SQL
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs3		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True : blnData3 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' --������ �о�´�.
		Call SubMakeSQLStatements("A",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("B",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		
		Call SubMakeSQLStatements("C",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
		
		
		
		If blnData1 = False And blnData2 = False And   blnData3 = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Call MoveFirst(pType)
		Select Case pType
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
			Case 3
				lgoRs3.Find pWhereSQL	
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
			Case 3
				lgoRs3.Filter = pWhereSQL	
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
			Case 3
				EOF = lgoRs3.EOF	
		End Select
	End Function
	
	Function MoveFirst(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
			Case 3
				lgoRs3.MoveFirst	
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
			Case 3
				lgoRs3.MoveNext	
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
		On Error Resume Next
		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
					If Err Then PrintLog "pFieldNm=" & pFieldNm : Reponse.End
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
				End If
			Case 3
				If Not lgoRs3.EOF Then
					GetData = lgoRs3(pFieldNm)
				End If	
		End Select
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)		
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf	
				lgStrSQL = lgStrSQL & " FROM TB_8_3_A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				lgStrSQL = lgStrSQL & "  ORDER BY  CAST(W101 as INT)  ASC " & vbcrlf

	      Case "B"
				If WHERE_SQL <> "" Then
				
					lgStrSQL = ""
					lgStrSQL = lgStrSQL & " SELECT  "
					lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
					lgStrSQL = lgStrSQL & " FROM TB_8_3_B	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
					lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				Else
				
					lgStrSQL =			  "SELECT "
					lgStrSQL = lgStrSQL & "  SEQ_NO   , A.W105, A.W105_NM, A.W106, A.C_W107, A.C_W108, A.C_W109, A.C_W110, A.C_W111, A.C_W112, A.C_W113, A.C_W114, A.C_W115, A.C_W116, A.C_W117, A.C_W118 "
					lgStrSQL = lgStrSQL & " FROM TB_8_3_B A WITH (NOLOCK) "
					lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & "  union all"
					lgStrSQL = lgStrSQL & "	 SELECT 9999999 SEQ_NO ,  '99999' , '��','',Sum(A.C_W107), Sum(A.C_W108), Sum(A.C_W109), Sum(A.C_W110), Sum(A.C_W111), Sum(A.C_W112), Sum(A.C_W113), Sum(A.C_W114), Sum(A.C_W115), Sum(A.C_W116),Sum(A.C_W117), Sum(A.C_W118 )"
					lgStrSQL = lgStrSQL & "  FROM TB_8_3_B A WITH (NOLOCK) "
					lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & "  ORDER BY  W105 ,  W106  ASC " & vbcrlf
					
					
				End If	

				
			Case "C"	
				 lgStrSQL = ""
				 lgStrSQL = lgStrSQL &  " SELECT Sum(A.C_W107) C_W107, Sum(A.C_W108) C_W108, Sum(A.C_W116) C_W116,Sum( A.C_W118)  C_W118  " & vbCrLf
				 lgStrSQL = lgStrSQL &  " FROM TB_8_3_B A WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				 lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
				 lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				 lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				 If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6121MA1
	Dim A103
	Dim A181
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6121MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo , dblSum , dblCode30
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6121MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6121MA1"

	Set lgcTB_8_3 = New C_TB_8_3		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_8_3.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W6121MA1

	'==========================================
	' --  ��8ȣ �������鼼�װ�꼭(3) �������� 
	' -- 1. ����׸��԰ŷ��� 
	
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	dblSum = 0
	dblCode30 = 0	
	
	lgcTB_8_3.Find 1, "W_Code='01'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='16'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='02'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='04'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='05'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='06'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='07'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='08'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='09'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='10'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='11'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='15'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='12'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='17'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='18'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='13'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "W101_NM"), 40)      '���� 
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='30'"
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='19'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	lgcTB_8_3.Find 1, "W_Code='20'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '��곻�� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
	
	sHTFBody = sHTFBody & UNIChar("", 29) & vbCrLf	' -- ���� 
	
	' -- �����Ǿ� ���ϻ��� ������ �ٲ� 
	' ---------------------------------------------------------

	lgcTB_8_3.Find 1, "W_Code='01'"
	
	Do Until lgcTB_8_3.EOF(1) 
		SELECT CASE  lgcTB_8_3.GetData(1, "W_CODE")
		
	
		CASE   "13" , "14"
		
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "W101_NM"), 40)      '���� 
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)          '��곻�� 
		    	'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '������󼼾� 
		
		       dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '�� ���������� �� 
		
		
		CASE   "30" 
		
		      '������� �հ� 
		
			If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W104"), lgcTB_8_3.GetData(1, "W101_NM") & "_������󼼾�") Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0)
		
		    dblCode30 = unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)  '������� �հ� 

' -- 2006.03.27 �����߰� 
'    ���װ�����������(3)(A149)������ �ڵ�04. �����η°��ߺ񼼾װ����� �׸�(104)������󼼾��� 0���� Ŭ ��� 
'     �������η°��ߺ�߻�����(A181)�� ������ ������ ���� 
'    -> ���� �߰� 
		Case "04"

			If UNICDbl(lgcTB_8_3.GetData(1, "C_W104"), 0) > 0 Then	' -- ��󼼾� 
			  '- �ڵ�(32)�� ����������"0"���� ū ��� �������η°��ߺ�߻�����(A181)�� �������� ���� 

				Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp �� ���ǵ� 
								
				' -- �߰� ��ȸ������ �о�´�.
				cDataExists.A181.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
				cDataExists.A181.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
				If Not cDataExists.A181.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "", lgcTB_8_3.GetData(1, "W101_NM") & "_��󼼾��� '0'���� ū ��� �������η°��ߺ�߻�����(A181) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
						
				End If
							
				' -- ����� Ŭ���� �޸� ���� 
				Set cDataExists.A181 = Nothing							
			End If	

			dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '�� ���������� �� 
		CASE ELSE
		      	'If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W103"), lgcTB_8_3.GetData(1, "W101_NM") & "_��곻��") Then blnError = True	' -- Null ��� 
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)
		
				If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W104"), lgcTB_8_3.GetData(1, "W101_NM") & "_������󼼾�") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0)
		
		       dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '�� ���������� �� 
		      
		END SELECT		
				
		
		Call lgcTB_8_3.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop
	lgcTB_8_3.WHERE_SQL = ""
	      
	
	if dblCode30 <> dblSum then
	    Call SaveHTFError(lgsPGM_ID, dblCode30 & " <> " & dblSum, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"������� �հ�", "�� ���������� ��"))
	    blnError = True	
	End if
	
	if Not lgcTB_8_3.EOF(3)  then
 
		if dblCode30 <> unicdbl(lgcTB_8_3.GetData(3, "C_W107"),0) then
		    Call SaveHTFError(lgsPGM_ID, dblCode30, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"������� �հ�", "�׸�(107)���������_������ �հ�"))
		    blnError = True	
		End if
	End if
	

	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	' -- 2. �����̿��װ�� 
	iSeqNo = 1

	Do Until lgcTB_8_3.EOF(2) 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ���			
		
		If UNICDbl(lgcTB_8_3.GetData(2, "SEQ_NO"), 0) <> 9999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "W106"), "�������") Then blnError = True	' �հ��� �ܿ� ������� �ʼ�üũ 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(2, "SEQ_NO"), 6)
		End If
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "W105"), "����") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(2, "W105"), 2)
		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_8_3.GetData(2, "W106"))	' ������� 

		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W107"), "����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W107"), 15, 0)
		
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W108"), "�̿���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W108"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W109"), "��������󼼾�_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W109"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W110"), "��������󼼾�_1��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W110"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W111"), "��������󼼾�_2��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W111"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W112"), "��������󼼾�_3��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W112"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W113"), "��������󼼾�_4��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W113"), 15, 0)
		
		If  ChkNotNull(lgcTB_8_3.GetData(2, "C_W114"), "��������󼼾�_�հ�") Then 
		    if unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W109"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W111"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W112"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W113"),0) then
		       Call SaveHTFError(lgsPGM_ID, unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") & "��������󼼾�_�հ�", "�׸�(109)+�׸�(110)+�׸�(111)+�׸�(112)+�׸�(113)"))
		       blnError = True	
		    End if
		  
		  
		Else
		    blnError = True	
		End if
		  sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W114"), 15, 0)     
		 
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W115"), "�����Ѽ����뿡 ���� �̰�����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W115"), 15, 0)
		
		If  ChkNotNull(lgcTB_8_3.GetData(2, "C_W116"), "��������") Then 
		    if unicdbl(lgcTB_8_3.GetData(2, "C_W116"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W115"),0)  then
		       Call SaveHTFError(lgsPGM_ID, lgcTB_8_3.GetData(2, "C_W116"), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") &"��������", "�׸�(114)-�׸�(115)"))
		       blnError = True	
		    End if
		Else
		    blnError = True	
		End if    
		  
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W116"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W117"), "�Ҹ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W117"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W118"), "�̿���") Then blnError = True	
		     if unicdbl(lgcTB_8_3.GetData(2, "C_W118"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W107"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W108"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W116"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W117"),0) then
		       Call SaveHTFError(lgsPGM_ID, unicdbl(lgcTB_8_3.GetData(2, "C_W118"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") &"�̿���", "�׸�(107)+�׸�(108) -�׸�(116)-�׸�(117) "))
		       blnError = True	
		 End if
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W118"), 15, 0)

		sHTFBody = sHTFBody & UNIChar("", 38) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_8_3.MoveNext(2)	' -- 2�� ���ڵ�� 
	Loop

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_8_3 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6121MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W6121MA1 : " & lgStrSQL
End Sub
%>
