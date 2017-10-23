<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��68ȣ �ұް������μ���ȯ�޽�û�� 
'*  3. Program ID           : W8105MA1
'*  4. Program Name         : W8105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_68

Set lgcTB_68 = Nothing	' -- �ʱ�ȭ 

Class C_TB_68
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1
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

		lgStrSQL = ""
		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgoRs1,lgStrSQL, "", "") = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		
		LoadData = True
	End Function

	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
	            
	            If WHERE_SQL = "" Then	' �ܺ�ȣ���� �ƴϸ� ���������� ������� �ҷ��´� 
					lgStrSQL = lgStrSQL & " , B.BANK_CD, B.BANK_BRANCH, B.BANK_DPST, B.BANK_ACCT_NO " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " FROM TB_68	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				
				If WHERE_SQL = "" Then	' �ܺ�ȣ���� �ƴϸ� ���������� �����Ѵ�.
					lgStrSQL = lgStrSQL & " INNER JOIN TB_COMPANY_HISTORY B WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W8105MA1
	Dim A101

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8105MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
  '  On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8105MA1"
	
	Set lgcTB_68 = New C_TB_68		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_68.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
	Set cDataExists = new  TYPE_DATA_EXIST_W8105MA1		
	'==========================================
	' -- ��68ȣ �ұް������μ���ȯ�޽�û�� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_68.GetData("W1_S"), "��ջ������_������") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W1_S"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W1_E"), "��ջ������_����") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W1_E"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W2_S"), "�����������_������") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W2_S"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W2_E"), "�����������_����") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W2_E"))
	
	If  ChkNotNull(lgcTB_68.GetData("W6"), "��ձݾ�") Then 
	    
		Set cDataExists.A101  = new C_TB_3	' -- W8101MA1_HTF.asp �� ���ǵ� 
		Call SubMakeSQLStatements_W8105MA1("0",iKey1, iKey2, iKey3)   
		'cDataExists.A101.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		'cDataExists.A101.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
		
		 If Not cDataExists.A101.LoadData()  then
			    blnError = True
				Call SaveHTFError(lgsPGM_ID, "�� 3ȣ ���μ�����ǥ�ع׼���������꼭(A101)", TYPE_DATA_NOT_FOUND)	
         Else

				 if    UNICDbl(cDataExists.A101.W06, 0) >= 0  and UNICDbl(lgcTB_68.GetData("W6"),0)  <> 0 then
				     Call SaveHTFError(lgsPGM_ID,lgcTB_68.GetData("W6"), UNIGetMesg("��û����� �ƴմϴ�","", ""))
				     blnError = True	
				 Else    
						'��ձݾ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(06)������⵵�ҵ�ݾ� X (-1)�� ���ƾ� �� 
						if   UNICDbl(lgcTB_68.GetData("W6"),0) <> UNICDbl(cDataExists.A101.W06, 0) * -1 then
						     Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��ձݾ�", "�� 3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(06)������⵵�ҵ�ݾ� X (-1)"))
						    blnError = True		
						end if
        
        
				end if
		end if		
			Set cDataExists.A101 = Nothing	
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W6"), 15, 0)
	
	'�ұް������� ��ձݾ׶��� (8)���� ����ǥ�رݾװ� (6)���� ��ձݾ׺��� �۰ų� ���ƾ� ��,
	If  ChkNotNull(lgcTB_68.GetData("W7"), "�ұް���������ձݾ�") Then
	    if UNICDbl(lgcTB_68.GetData("W7"),0) > UNICDbl(lgcTB_68.GetData("W8") ,0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W8"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "�ұް���������ձݾ�", "����ǥ��"))
	        blnError = True		
	    end if
	    
	    if UNICDbl(lgcTB_68.GetData("W7"),0) > UNICDbl(lgcTB_68.GetData("W6") ,0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W6"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "�ұް���������ձݾ�", "��ձݾ�"))
	        blnError = True		
	    end if
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W7"), 15, 0)
	
	

	
	If Not ChkNotNull(lgcTB_68.GetData("W8"), "����ǥ��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W8"), 15, 0)
	
	If Not ChkNotNull(Replace(lgcTB_68.GetData("W9"),"%",""), "����") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(Replace(lgcTB_68.GetData("W9"),"%",""), 15, 0)
	
	If Not ChkNotNull(lgcTB_68.GetData("W10"), "���⼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W10"), 15, 0)
	
	If Not ChkNotNull(lgcTB_68.GetData("W11"), "�������鼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W11"), 15, 0)
	
	If ChkNotNull(lgcTB_68.GetData("W12"), "��������") Then
	   '�������� = ���⼼�� - �������鼼�� 
	   if   UNICDbl(lgcTB_68.GetData("W12"),0) <>  UNICDbl(lgcTB_68.GetData("W10"),0) - UNICDbl(lgcTB_68.GetData("W11"),0) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W12"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������", "���⼼�� - �������鼼��"))
	        blnError = True		
	   End if     
	Else 
		  blnError = True		
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W12"), 15, 0)

	
	If  ChkNotNull(lgcTB_68.GetData("W13"), "��������������μ���") Then
	    If UNICDbl(lgcTB_68.GetData("W13"),0) <> UNICDbl(lgcTB_68.GetData("W10"),0) then
	       '��������������μ���: �׸� (10)���⼼�װ� ���� 
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������������μ���", "���⼼��"))
	       blnError = True		
	    End if
	Else
	      blnError = True		
	End if    
	    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W13"), 15, 0)
	'�����Ҽ��� : �����Ҽ��� �� (���⼼�� - ��������)
	If  ChkNotNull(lgcTB_68.GetData("W14"), "�����Ҽ���") Then
	    if UNICDbl(lgcTB_68.GetData("W14"),0) <  UNICDbl(lgcTB_68.GetData("W10"),0) - UNICDbl(lgcTB_68.GetData("W12"),0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W14"), UNIGetMesg(TYPE_CHK_OVER_EQUAL, "�����Ҽ���", "(���⼼�� - ��������)"))
	   	        blnError = True	
	    End if
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W14"), 15, 0)
	
	
	'ȯ�޽�û���� : ��������������μ��� - �����Ҽ��� 
	If  ChkNotNull(lgcTB_68.GetData("W15"), "ȯ�޽�û����") Then 
	    if UNICDbl(lgcTB_68.GetData("W15"),0) <> UNICDbl(lgcTB_68.GetData("W13"),0) - UNICDbl(lgcTB_68.GetData("W14"),0) Then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ȯ�޽�û����", "��������������μ��� - �����Ҽ���"))
	        blnError = True		
	    End if
	    
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W15"), 15, 0)
	
	if (Trim(lgcTB_68.GetData("BANK_CD")) = "" and Trim(lgcTB_68.GetData("BANK_ACCT_NO")) <> "")  Or (Trim(lgcTB_68.GetData("BANK_CD")) <> "" and Trim(lgcTB_68.GetData("BANK_ACCT_NO")) = "") Then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("BANK_CD") & Trim(lgcTB_68.GetData("BANK_ACCT_NO")), UNIGetMesg("�����ڵ� �Ǵ� ���¹�ȣ�� �� �� �ϳ��� �ԷµǾ� �ֽ��ϴ�", "", ""))
	        blnError = True		
	End if
	
	' Null ��� 
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_CD"), "����ó(����)�ڵ�") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_CD"), 2)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_BRANCH"), "����ó(��)����") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_BRANCH"), 20)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_DPST"), "��������") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_DPST"), 20)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_ACCT_NO"), "���¹�ȣ") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_ACCT_NO"), 20)
	

	sHTFBody = sHTFBody & UNIChar("", 60)	' -- ���� 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_68 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_W8105MA1 : " & lgStrSQL
End Sub

%>
