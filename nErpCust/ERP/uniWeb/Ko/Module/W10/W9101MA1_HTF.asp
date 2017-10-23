<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��47ȣ �ֿ��������(��)
'*  3. Program ID           : W9101MA1
'*  4. Program Name         : W9101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_47A

Set lgcTB_47A = Nothing ' -- �ʱ�ȭ 

Class C_TB_47A
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W2_CD
	Dim W3
	Dim W4
	Dim W5
	
	Dim W124
	Dim W125
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs3
				 
		On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' --������ �о�´�.
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

		' -- ������ �о�´�.
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3) 
		                                      '�� : Make sql statements
		If   FncOpenRs("R",lgObjConn,oRs3,lgStrSQL, "", "") = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		'Response.End 'zzz
		W124		= oRs3("W124")
		W125		= oRs3("W125")
		

		
		Call SubCloseRs(oRs3)
		
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
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_CD		= lgoRs1("W2_CD")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
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
				lgStrSQL = lgStrSQL & " FROM TB_47A1 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				'lgStrSQL = lgStrSQL & "		AND A.W2_CD<>'75' " & pCode3 	' 200703 TEMP
				
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_47A2 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

				
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9101MA1
	Dim A129
	DIM A115
	Dim A101
	Dim A137
	Dim A142
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists,sTmp2
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9101MA1"

	Set lgcTB_47A = New C_TB_47A		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_47A.LoadData Then Exit Function			
	
	Set cDataExists = new TYPE_DATA_EXIST_W9101MA1
	'==========================================
	' -- ��47ȣ �ֿ��������(��) �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
 
	Do Until lgcTB_47A.EOF 
	

		Select Case lgcTB_47A.W2_CD 
		

		   	Case "41"
		   	   '�ڵ�(41)�翬�ձݱ�α��� �׸�(3)ȸ����ݾ�	- ��αݸ���(A129)�� ������α�(10), ��ġ�ڱ�(20), ��ȭ����(60)�� �հ� ��ġ 
			   '(�ڵ�(41)�� �׸�(3)�� �ݾ��� ��0������ ū ��� A129 �ݵ�� �Է�)
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
			    
		 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_ȸ����ݾ�(41) '0'���� ū ��� ��αݸ���(A129) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '10' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '20' " 
							     sTmp = sTmp +  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '60' " 
							     sTmp = sTmp +  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(41)�翬�ձݱ�α��� �׸�(3)ȸ����ݾ�"," ��αݸ���(A129)�� ������α�(10), ��ġ�ڱ�(20), ��ȭ����(60)�� ��"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
				
				
				
		     Case "64"
		   	  '�ڵ�(64)50%�ձݱ�α��� �׸�(3)ȸ����ݾ�	- ��αݸ���(A129)�� ��α�(30)�� �հ� ��ġ 
		   	  '(�ڵ�(64)�� �׸�(3)�� �ݾ��� ��0������ ū ��� A129 �ݵ�� �Է�)
		
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
						
								Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W2, UNIGetMesg(lgcTB_47A.W2 & "_ȸ����ݾ� '0'���� ū ��� ��αݸ���(A129) ���� �ʼ� �Է�", "",""))
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '30' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "�ڵ�(64)50%�ձݱ�α��� �׸�(3)ȸ����ݾ�","  ��αݸ���(A129)�� ��α�(30)�� �հ� ��ġ"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
				
				
		  Case "42"
		   	  '�ڵ�(42)������α��� �׸�(3)ȸ����ݾ�	- ��αݸ���(A129)�� ������α�(40), ��ȭ��ü(70)�� �հ� ��ġ 
		   	  '(�ڵ�(42)�� �׸�(3)�� �ݾ��� ��0������ ū ��� A129 �ݵ�� �Է�)
		   	  
		
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_ȸ����ݾ� '0'���� ū ��� ��αݸ���(A129) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '40' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '70' " 
							     sTmp = sTmp + UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "�ڵ�(42)������α��� �׸�(3)ȸ����ݾ�","  ��αݸ���(A129)�� ��Ÿ��α��� �׸�(50) �հ� ��ġ"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
				
				
				
			Case "73"
			   ' �ڵ�(73)��Ÿ��α��� �׸�(3)ȸ����ݾ��ڵ�(73)��Ÿ��α��� �׸�(4)���������(����)�ݾ�	
			    '- ��αݸ���(A129)�� ��,��Ÿ��α��� �׸�(50)�� ��ġ 
			    '(�ڵ�(73)�� �׸�(3),�ڵ�(73)�� �׸�(4)�� �ݾ��� ��0������ ū ��� A129 �ݵ�� �Է�)
				
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_ȸ����ݾ� '0'���� ū ��� ��αݸ���(A129) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '50' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "��Ÿ��α��� �׸�(3)ȸ����ݾ�","  ��αݸ���(A129)�� ������α�(40), ��ȭ��ü(70)�� �հ� ��ġ"))
								End If
								
								If UNICDbl(lgcTB_47A.W4, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "��Ÿ��α��� �׸�(4)���������(����)�ݾ�","  ��αݸ���(A129)�� ������α�(40), ��ȭ��ü(70)�� �հ� ��ġ"))
								End If
								
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
				
				
				
				
				
			Case "65"
			   '�ڵ�(65)�����(5�����ʰ�)�� �׸�(3)ȸ����ݾ�	
				'- ǥ�ؼ��Ͱ�꼭 �� �μӸ����� �����ݾ��� �ִµ� �ڵ�(65)�� �׸�(3) �� �ݾ��� ������ ���� 
				'(A115 �ڵ�(25), A116 �ڵ�(37), A117 �ڵ�(27), A118 �ڵ�(22), A119 �ڵ�(14), A120 �ڵ�(24), A121 �ڵ�(27), A123 �ڵ�(25))
				
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
			      
			        	Set cDataExists.A115 = new C_TB_3_3	' -- W5109MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A115.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A115.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A115.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "ǥ�ؼ��Ͱ�꼭 �� �μӸ����� �����ݾ��� �ִµ� �ڵ�(65)�� �׸�(3) �� �ݾ��� ������ ���� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
							
							    
								IF cDataExists.A115.W1 = 1 THEN 
								   cDataExists.A115.Find "w4 = 25 "
								ELSE
								   cDataExists.A115.Find "w4 = 37 "
								END IF      
								 
								sTmp =  UNICDbl(cDataExists.A115.W5 ,0)
									If UNICDbl(sTmp , 0) > 50000  and UNICDbl(lgcTB_47A.W3,0) = 0 Then
										blnError = True
										Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg("ǥ�ؼ��Ͱ�꼭�� ����� �׸� �ݾ��� ������ �ڵ�(65)�����(5�����ʰ�)�׸� �ݾ��� �ԷµǾ���մϴ�.","",""))
									End If
								
								
								
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A115 = Nothing				
			
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)	
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
				
				
				
			Case "66"
			
			    If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then blnError = True	
  		        sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)	
				
				
				
			   '�ڵ�(66)������α��ѵ����� �׸�(5)�������ݾ� 
			   '- ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ���� 100�������� ũ��, 
			   '��αݸ���(A129)�� ������α�_�谡 100���� �̻��� ��� ��0������ ū �� �Է� 
				
			    If  ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then 
						Set cDataExists.A101 = new C_TB_3	' -- W1109MA1_HTF.asp �� ���ǵ� 
									
							' -- �߰� ��ȸ������ �о�´�.
							'cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							'cDataExists.A101.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
									
							If Not cDataExists.A101.LoadData() Then
								blnError = True
								'Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_�������ݾ� '0'���� ū ��� ǥ�ؼ��Ͱ�꼭 ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							Else
								
								'sTmp = UNICDBL(cDataExists.A101.W10  ,0)
								sTmp = UNICDBL(cDataExists.A101.W56  ,0)	' -- 2006-01-05 : 200603 ������ 
							
							End If
					
								
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A101 = Nothing		
			    
			         if sTmp > 1000000 Then 
			    
							Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp �� ���ǵ� 
											
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A129.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A129.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
							If Not cDataExists.A129.LoadData() Then
								'blnError = True
								'Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg( lgcTB_47A.W2 & "_ȸ����ݾ� '0'���� ū ��� ��αݸ���(A129) ���� �ʼ� �Է� ","",""))		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
								sTmp = 0
							Else
									    
							     cDataExists.A129.Find 2, "w9_cd = '99' " 
							     sTmp2 =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
										
							End If
								
										
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A129 = Nothing
					
							If sTmp2 > 1000000 and unicdbl(lgcTB_47A.W5,0)  <= 0 then
					
							     Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg( lgcTB_47A.W2 & "_�������ݾ��� ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ����_ 100�������� ũ�� ��αݸ���(A129)�� ������α�_�谡 100���� �̻��� ��� ��0������ ū ���� �Է� " ,"",""))		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
							     blnError = True
							End If		
					End if					
			Else
			   	blnError = True
			End if		
			
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
			
				
			
			Case "49", "75", "76"
			
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�����ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)

			Case "47"
				If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
						
			Case Else
			
				If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "ȸ����ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "���������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "�������ݾ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
		End Select
		
			
	
		'�׸�(5)�������ݾ� =  �׸�(3)ȸ����ݾ� - �׸�(4)���������(����)�ݾ�(�ڵ� 53,12,71,13,72,55,56,57,58,59,41,64,42,65,61,74,77)
		SELECT CASE lgcTB_47A.W2_CD
		  CASE 53,12,71,13,72,55,56,57,58,59,41,64,42,65,61,74,77
				If UNICDbl(lgcTB_47A.W5,0)  <> UNICDbl(lgcTB_47A.W3,0) - UNICDbl(lgcTB_47A.W4,0) Then
				   Call SaveHTFError(lgsPGM_ID,lgcTB_47A.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_47A.W2 &  "�������ݾ�","ȸ����ݾ�-���������ݾ�"))
				   blnError = True	
				End if
		END SELECT  
		
		lgcTB_47A.MoveNext 
	Loop

	If  ChkNotNull(lgcTB_47A.W124, "�󿩹���_�ҵ�ó�бݾ�") Then 
	
		'�ڵ�(97)�ҵ�ó�бݾ�	- �ҵ��ڷ��(A137)�� �׸�(4)�ҵ�ݾ�_��� ��ġ(�ڵ�(97)�� ��0������ ū ��� A137 �ݵ�� �Է�)
			
				    
		if UNICDbl(lgcTB_47A.W124,0) > 0  Then		    
				    
			Set cDataExists.A137 = new C_TB_55	' -- W5109MA1_HTF.asp �� ���ǵ� 
											
			' -- �߰� ��ȸ������ �о�´�.
			cDataExists.A137.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A137.WHERE_SQL = "and SEQ_NO = '999999' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
			If Not cDataExists.A137.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W124, "�ҵ��ڷ��(A137) ������ �ۼ����� �ʾҽ��ϴ�.")
			Else
			
			     sTmp =  UNICDbl(cDataExists.A137.GetData("W4"),0)
							
											
			End If
								
										
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A137 = Nothing
						
			If  sTmp <> UNICDbl(lgcTB_47A.W124,0) then
			  	blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W124 & " <> " & sTmp, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"�ڵ�(97)�ҵ�ó�бݾ�","  �ҵ��ڷ��(A137)�� �׸�(4)�ҵ�ݾ�_��"))
			End If	
		End If					
	Else
	   	blnError = True
	End if		
			
	'200703
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '�Һ񼺼��񽺾���������_ȸ����ݾ� 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '�Һ񼺼��񽺾���������_���������ݾ� 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '�Һ񼺼��񽺾���������_�������ݾ� 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '��Ÿ�� ��������_ȸ����ݾ� 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '��Ÿ�� ��������_�������ݾ� 
		
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W124, 15, 0)
						
	If Not ChkNotNull(lgcTB_47A.W125, "�󿩹���_����ó�б�") Then blnError = True	

	' -- 2006.03.24�߰� : ���α��� '2'�� �������� 
	if UNICDbl(lgcTB_47A.W125,0) > 0 And lgcCompanyInfo.Comp_type2 <> "2" Then		    
				    
		Set cDataExists.A142 = new C_TB_3_3_4	' -- W5109MA1_HTF.asp �� ���ǵ� 
											
		' -- �߰� ��ȸ������ �о�´�.
		cDataExists.A142.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A142.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
		If Not cDataExists.A142.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W125, "�����׿���ó��(���ó��)��꼭(A142) ������ �ۼ����� �ʾҽ��ϴ�.")
		Else
			
		     sTmp =  UNICDbl(cDataExists.A142.W5,0) + UNICDbl(cDataExists.A142.W15,0) + UNICDbl(cDataExists.A142.W26,0)
							
											
		End If
								
										
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A142 = Nothing
						
		If  sTmp <> UNICDbl(lgcTB_47A.W125,0) then
		  	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W125 & " <> " & sTmp, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"�ڵ�(98)��.���� ����ó�б�","  �����׿���ó��(���ó��)��꼭(A142)�� �ڵ�(5)�߰����� + �ڵ�(15)���� + �ڵ�(26)����ó�п����ѻ󿩱�"))
		End If	
	End If					


	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W125, 15, 0)
				
	sHTFBody = sHTFBody & UNIChar("", 49)	' -- ���� 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_47A = Nothing	' -- �޸����� 
	
End Function


%>
