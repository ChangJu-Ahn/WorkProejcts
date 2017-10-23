<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��2ȣ �����Ư���� ����ǥ�� �� ���׽Ű� 
'*  3. Program ID           : W8113MA1
'*  4. Program Name         : W8113MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_2

Set lgcTB_2 = Nothing	' -- �ʱ�ȭ 

Class C_TB_2
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W7
	Dim W8
	Dim W9
	Dim W10_1
	Dim W10_2
			
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

		lgStrSQL = ""
		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W1				= oRs1("W1")
		W2				= oRs1("W2")
		W3				= oRs1("W3")
		W4				= oRs1("W4")
		W5				= oRs1("W5")
		W6				= oRs1("W6")
		W7				= oRs1("W7")
		W8				= oRs1("W8")
		W9				= oRs1("W9")
		W10_1			= oRs1("W10_1")
		W10_2			= oRs1("W10_2")
		
		Call SubCloseRs(oRs1)	
		
		LoadData = True
	End Function

	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub	

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_2	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W8113MA1
	Dim A105
	Dim A101
	
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8113MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8113MA1"
	
	Set lgcTB_2 = New C_TB_2		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_2.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
	
	Set cDataExists = new TYPE_DATA_EXIST_W8113MA1	
	'==========================================
	' -- 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	

	
	If ChkNotNull(lgcTB_2.W1, "����ǥ��")  Then ' -- ����Ÿ����� ������ 
	
			
			' -- ��12ȣ�����Ư��������ǥ�ع׼���������꼭(A105)
			Set cDataExists.A105  = new C_TB_12	' -- W8111MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8111MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A105.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A105.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A105.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��12ȣ�����Ư��������ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				'�Ϲݹ��� : �����Ư��������ǥ�ع׼���������꼭(A105)�� �׸�(8)  �Ϲݹ��� ����ǥ�رݾ�_�Ұ�Ǵ� 
				'���չ��� : �����Ư��������ǥ�ع׼���������꼭(A105)�� �׸�(12) ���չ��� ����ǥ�رݾ�_�Ұ�� ��ġġ���� ������ ���� 
				If (UNICDbl(lgcTB_2.W1, 0) <> UNICDbl(cDataExists.A105.w8_Amt, 0))  And  (UNICDbl(lgcTB_2.W1, 0) <> UNICDbl(cDataExists.A105.w12_Amt, 0) )Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ��(7)","��12ȣ�����Ư��������ǥ�ع׼���������꼭(A105) �׸�(8)  ����ǥ�رݾ�_�Ұ�"))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A105 = Nothing
		
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W1, 15, 0)
	

	If ChkNotNull(lgcTB_2.W2, "���⼼��")  Then ' -- ����Ÿ����� ������ 
	
			
			' -- ��12ȣ�����Ư��������ǥ�ع׼���������꼭(A105)
			Set cDataExists.A105  = new C_TB_12	' -- W8111MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A105.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A105.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A105.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��12ȣ�����Ư��������ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				'�Ϲݹ��� : �����Ư��������ǥ�ع׼���������꼭(A105)�� �׸�(8)  �Ϲݹ��� ���⼼��_�Ұ�Ǵ� 
				'���չ��� : �����Ư��������ǥ�ع׼���������꼭(A105)�� �׸�(12) ���չ��� ���⼼��_�Ұ�� ��ġġ���� ������ ���� 
				If (UNICDbl(lgcTB_2.W2, 0) <> UNICDbl(cDataExists.A105.w8_Tax, 0))  And (UNICDbl(lgcTB_2.W2, 0) <> UNICDbl(cDataExists.A105.w12_Tax, 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ��(7)","�׸�(8)  ���⼼��_�Ұ�"))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A105 = Nothing
	
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W2, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W3, "���꼼��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W3, 15, 0)
	
	'�Ѻδ㼼��(10) : (8)���⼼�� + (9)���꼼�� 
	
	If ChkNotNull(lgcTB_2.W4, "�Ѻδ㼼��") Then
	    if UNICDbl(lgcTB_2.W4, 0) <> UNICDbl(lgcTB_2.W2,0) + UNICDbl(lgcTB_2.W3, 0) then
	       	blnError = True
		
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ѻδ㼼��","���⼼�� + ���꼼��"))
	    end if
	else
	   blnError = True	
	end if  

	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W4, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W5, "�ⳳ�μ���") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W5, 15, 0)
	
	
	'���������Ҽ���(12) : (10)�Ѻδ㼼�� - (11)�ⳳ�μ��� 
	If ChkNotNull(lgcTB_2.W6, "���������Ҽ���") Then
	    if UNICDbl(lgcTB_2.W6, 0) <> UNICDbl(lgcTB_2.W4,0) - UNICDbl(lgcTB_2.W5, 0) then
	       	blnError = True
		
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���������Ҽ���","�Ѻδ㼼�� - �ⳳ�μ���"))
	    end if
	else
	   blnError = True	
	end if  	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W6, 15, 0)
	
	
	'�������μ���(14)  : (12)���������Ҽ��� - (13)�г��Ҽ��� 
	If Not ChkNotNull(lgcTB_2.W7, "�г��Ҽ���") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W7, 15, 0)
	

	If ChkNotNull(lgcTB_2.W8, "�������μ���") Then
	    if UNICDbl(lgcTB_2.W8, 0) <> UNICDbl(lgcTB_2.W6,0) - UNICDbl(lgcTB_2.W7, 0) then
	       	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������μ���","���������Ҽ��� - �г��Ҽ���"))
	    end if
	else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W8, 15, 0)
	

	'����ĳ��μ���(15) : (14)�������μ���  - (16)����ҳ����Ư���� 
	If ChkNotNull(lgcTB_2.W9, "����ĳ��μ���") Then
	    if UNICDbl(lgcTB_2.W9, 0) <> UNICDbl(lgcTB_2.W8,0) - UNICDbl(lgcTB_2.W10_2, 0) then
	       	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W9, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ĳ��μ���","�������μ��� - ����ҳ����Ư����"))
	    end if
	else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W9, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W10_1, "����ȯ�ޱ�����û_ȯ�޹��μ�") Then blnError = True	
	
	'- ������ ���� 
	'- ���μ��� ȯ���� ��� �Է°��� 
	'- �ԷµȰ�� ZERO���� ũ�� ���μ�����ǥ�ع׼���������꼭(A101)�� 
	'  �ڵ�(46)���������Ҽ��װ�(�ڵ�(46) >= 0 ����)�� ���밪�� ��ġ�ؾ� �� 
    '  (��: �ڵ�(46)�� -100,000 �̸� ȯ�޹��μ��� ZERO�̰ų� 100,000 ���� �ԷµǾ�� ��)
	'- ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(46)���������Ҽ��װ� >= 0 �ԷºҰ� 
	
	

	If UNICDbl(lgcTB_2.W10_1, 0) >= 0 Then	' -- ȯ�޹��μ� ���� ����üũ 
	   	if  UNICDbl(lgcTB_2.W10_1, 0) > 0 then	
				' -- ��3ȣ���μ�����ǥ�ع׼���������꼭(A101)
				Set cDataExists.A101  = new C_TB_3	' -- W8111MA1_HTF.asp �� ���ǵ� 
			
				' -- �߰� ��ȸ������ �о�´�.
				Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
				cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
				cDataExists.A101.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
				If Not cDataExists.A101.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "��3ȣ���μ�����ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
				Else
		
					If  (abs(UNICDbl(lgcTB_2.W10_1, 0) <> UNICDbl(cDataExists.A101.w46, 0)) Or UNICDbl(cDataExists.A101.w46, 0) >= 0)Then
						blnError = True
					
						Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ȯ�޹��μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(46)���������Ҽ��װ�(�ڵ�(46) >= 0 ����)�� ���밪"))
					End If
				End If
		
				' -- ����� Ŭ���� �޸� ���� 
				Set cDataExists.A101 = Nothing
		end if		
	
	
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_1, UNIGetMesg(TYPE_CHK_ZERO_OVER, "����ȯ�ޱ�����û_ȯ�޹��μ�",""))
	End If
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W10_1, 15, 0)
	
	
		
	If Not ChkNotNull(lgcTB_2.W10_2, "����ȯ�ޱ�����û_����ҳ����Ư����") Then blnError = True	
	'- ������ ���� 
	'- ���μ��� ȯ���� ��� �Է°��� 
	'- ZERO �̰ų� �ԷµȰ�� ZERO���� ũ�� ���μ�����ǥ�ع� ����������꼭(A101)
	'  �ڵ�(46)���������Ҽ��װ躸�� ���밪�� �۰ų� ���ƾ� �� 
	'- �ԷµȰ�� ZERO���� ũ�� �׸�(14)�������μ��׺��� �۰ų� ���ƾ� �� 
	
	If UNICDbl(lgcTB_2.W10_2, 0) >= 0 Then	' -- ȯ�޹��μ� ���� ����üũ 
	   		
			' -- ��3ȣ���μ�����ǥ�ع׼���������꼭(A101)
			Set cDataExists.A101  = new C_TB_3	' -- W8111MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A101.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A101.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��3ȣ���μ�����ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				
				If UNICDbl(lgcTB_2.W10_2, 0) > abs(UNICDbl(cDataExists.A101.w46, 0)) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_2, UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "ȯ�޹��μ�","���μ�����ǥ�ع� ����������꼭(A101)(46)���������Ҽ��װ躸�� ���밪�� "))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A101 = Nothing
	
	
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_2, UNIGetMesg(TYPE_CHK_ZERO_OVER, "����ȯ�ޱ�����û_����ҳ����Ư����",""))
	End If
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W10_2, 15, 0)


	sHTFBody = sHTFBody & UNIChar("", 29)	' -- ���� 
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If
	
	'Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_2 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 

	  Case "O" '-- �ܺ� ���� �ݾ� 
	
			
	End Select
	PrintLog "SubMakeSQLStatements_W8113MA1 : " & lgStrSQL
End Sub

%>
