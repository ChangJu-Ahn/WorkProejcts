<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��3ȣ��3(3) �μӸ� �������� 
'*  3. Program ID           : W1107MA1
'*  4. Program Name         : W1107MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_3_3_3

Set lgcTB_3_3_3 = Nothing ' -- �ʱ�ȭ 

Class C_TB_3_3_3
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
				'Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
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
		 Call GetData()
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
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
		End If
	End Function
	
	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
	
	Function Clone(Byref pRs)
		Set pRs = lgoRs1.clone
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
				lgStrSQL = lgStrSQL & " FROM TB_3_3_3	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W1107MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W1107MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim dblAmt1, dblAmt2 , dblAmt3, arrNew(50)	' -- ������ �ڵ� 
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1107MA1"

	Set lgcTB_3_3_3 = New C_TB_3_3_3		' -- �ش缭�� Ŭ���� 
	
	lgcTB_3_3_3.WHERE_SQL = "		AND A.W1 = '3' "		' 
	
	If Not lgcTB_3_3_3.LoadData Then Exit Function			' -- ��3ȣ��3(3) �μӸ� �������� ���� �ε� 
	
	
	'==========================================
	' -- ��3ȣ��3(3) �μӸ� �������� ���ڽŰ� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	Call lgcTB_3_3_3.Clone(oRs2)	' ���İ����� �ʿ��� ���� ���ڵ���� ���� 

	Do Until lgcTB_3_3_3.EOF 
	
		If  ChkNotNull(lgcTB_3_3_3.W5, lgcTB_3_3_3.W3) Then 
	    

					If lgcTB_3_3_3.W4 = "01" Then   '���� =�ڵ� 02 + 03 - 04
						oRs2.Find "W4 = '02'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						oRs2.Find "W4 = '03'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
					
						oRs2.Find "W4 = '04'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 - dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����","�ڵ� 02 + 03 - 04"))
						   blnError = True	
						End If

					End If
		
		
					If lgcTB_3_3_3.W4 = "05" Then	 '�빫�� =�ڵ� 06 + 07 + 08
						oRs2.Find "W4 = '06'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '07'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '08'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 + dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�빫��","�ڵ� 06 + 07 + 08"))
						   blnError = True	
						End If

					End If
		 
		
		
		
		
		
					If lgcTB_3_3_3.W4 = "09" Then   ' ��� = 10 + 11 + 12 + 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 22+ 23 + 24 + 25 + 26 + 27 + 28
						oRs2.Find "W4 = '10'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '11'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '12'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '13'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '14'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '15'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '16'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '17'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '18'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '19'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '20'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '21'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
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

						' -- 2006.03 ���� 
						oRs2.MoveFirst
						oRs2.Find "W4 = '35'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���","�ڵ� 10 + 11 + 12 + 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 35 + 22+ 23 + 24 + 25 + 26 + 27 + 28"))
						   blnError = True	
						End If	
						
			
					End If
		
		
					If lgcTB_3_3_3.W4 = "29" Then	 '�ڵ�(29)������������ = �ڵ� 01 + 05 + 09
						oRs2.MoveFirst				' ���� ���ڵ�� ���ư��� �˻��ϰ����Ҷ� MoveFirst�ؾߵ� 
						oRs2.Find "W4 = '01'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '05'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '09'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 + dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������������","�ڵ� 01 + 05 + 09"))
							blnError = True	
						End If

					End If

				
					If lgcTB_3_3_3.W4 = "31" Then	 '�ڵ�(31)�հ�= �ڵ� 29 + 30
						oRs2.Find "W4 = '29'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '30'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2  Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(31)�հ�"," �ڵ� 29 + 30"))
							blnError = True	
						End If
					End If
		
		
					If lgcTB_3_3_3.W4 = "34" Then	 '�ڵ�(34)�����ǰ�������� = �ڵ� 31 - 32 - 33
						oRs2.Find "W4 = '31'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '32'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '33'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)

						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 - dblAmt2 - dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(34)�����ǰ��������"," �ڵ� 31 - 32 - 33"))
							blnError = True	
						End If

					End If
						
		Else
		       blnError = True	
		End if
		
		' -- 2006.03 ���� 
		Select Case lgcTB_3_3_3.W4
			Case "35"
				arrNew(35) = lgcTB_3_3_3.W5
			Case Else
				sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_3.W5, 15, 0)
		End Select
		
		lgcTB_3_3_3.MoveNext 

	Loop
	
	' -- 2006.03 �������� 
	sHTFBody = sHTFBody & UNINumeric(arrNew(35), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 19)	' -- ����	2006.03���� 

	' ----------- 
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3_3_3 = Nothing	' -- �޸����� 

End Function


' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W1107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A117" '-- �ܺ� ���� SQL
		lgStrSQL = "		AND A.W1 = '3' "	 & vbCrLf
				
	End Select
	PrintLog "SubMakeSQLStatements_W1107MA1 : " & lgStrSQL
End Sub
%>
