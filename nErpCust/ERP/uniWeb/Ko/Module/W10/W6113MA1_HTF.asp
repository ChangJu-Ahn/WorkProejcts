<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��Ư��11ȣ��5�������Ư�����װ��� 
'*  3. Program ID           : W6113MA1
'*  4. Program Name         : W6113MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_JT11_5

Set lgcTB_JT11_5 = Nothing	' -- �ʱ�ȭ 

Class C_TB_JT11_5
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

	            If WHERE_SQL = "" Then	' �ܺ�ȣ���� �ƴϸ� ���������� ���������(â����) �ҷ��´� 
					lgStrSQL = lgStrSQL & " , B.FOUNDATION_DT " & vbCrLf
	            End If
	            	            
				lgStrSQL = lgStrSQL & " FROM TB_JT11_5	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 

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
Class TYPE_DATA_EXIST_W6113MA1
	Dim A165
	
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6113MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6113MA1"
	
	Set lgcTB_JT11_5 = New C_TB_JT11_5		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_JT11_5.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
	Set cDataExists = new  TYPE_DATA_EXIST_W6113MA1		
	'==========================================
	' -- ��Ư��11ȣ��5�������Ư�����װ��� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("FOUNDATION_DT"), "â��(�պ���)��") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT11_5.GetData("FOUNDATION_DT"))
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W1"), "������뼼�װ�����") Then 
	    '������뼼�װ����� : �׸�(7)��������ο��� x 1,000,000�� 
	    if unicdbl(lgcTB_JT11_5.GetData("W1"),0) <> Unicdbl(lgcTB_JT11_5.GetData("W2"),0) *  Unicdbl(lgcTB_JT11_5.GetData("W1_RATE_VALUE"),0)    then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������뼼�װ�����", "�׸�(7)��������ο��� x " & Unicdbl(lgcTB_JT11_5.GetData("W1_RATE_VALUE"),0)))
			blnError = True		
	    end if
	else
		blnError = True		
	end if	
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W1"), 15, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W2"), "��������ο���") Then
        '�׸�(8)���ذ���������ñٷ��ڼ� - �׸�(9)��������������ñٷ��ڼ�	     
      
	    if  unicdbl(lgcTB_JT11_5.GetData("W2"),0) <> fix(unicdbl(lgcTB_JT11_5.GetData("W3"),0) - unicdbl(lgcTB_JT11_5.GetData("W4"),0)) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W2"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������ο���", "�׸�(8)���ذ���������ñٷ��ڼ� - �׸�(9)��������������ñٷ��ڼ�"))
			blnError = True		
	    end if
	   
	else
	 blnError = True		
	End if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W2"), 5, 0)
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("W3"), "���ذ���������ñٷ��ڼ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W3"), 7, 2)
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("W4"), "��������������ñٷ��ڼ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W4"), 7, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W5"), "����������װ�����") Then 
	
	    '�׸�(11)��������ο��� x 500,000�� 
	    if unicdbl(lgcTB_JT11_5.GetData("W5"),0) <> fix(unicdbl(lgcTB_JT11_5.GetData("W6"),0) *  Unicdbl(lgcTB_JT11_5.GetData("W5_RATE_VALUE"),0)) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����������װ�����", "��������ο���  x " & Unicdbl(lgcTB_JT11_5.GetData("W5_RATE_VALUE"),0)))
	       blnError = True		
	    end if 
	Else
	    blnError = True		
	End if     
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W5"), 15, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W6"), "��������ο���") Then
	  
	  '��������ο��� : [(����������� ������ ��1������ñٷ���1�δ� 1����ձٷνð� 
	  '				-����������������� ��1������ñٷ���1�δ� 1����ձٷνð�)
	  '				/ ����������� ������ ��1������ñٷ���1�δ� 1����ձٷνð�]
	  '				x ��������������ñٷ��ڼ�  
      '              (�Ҽ����̸�����)
          if unicdbl(lgcTB_JT11_5.GetData("W8"),0)  <> 0 then
				if  unicdbl(lgcTB_JT11_5.GetData("W6"),0) <> fix(((unicdbl(lgcTB_JT11_5.GetData("W7"),0) - unicdbl(lgcTB_JT11_5.GetData("W8"),0))/unicdbl(lgcTB_JT11_5.GetData("W7"),0)) * unicdbl(lgcTB_JT11_5.GetData("W9"),0)) then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������ο���", "[(����������� ������ ��1������ñٷ���1�δ� 1����ձٷν�-����������������� ��1������ñٷ���1�δ� 1����ձٷνð�)/ ����������� ������ ��1������ñٷ���1�δ� 1����ձٷνð�]x ��������������ñٷ��ڼ�"))
				    blnError = True		
				end if
		  End if		

	   
	Else
	  blnError = True		
	end if
	
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W6"), 5, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W7"), "����������� �������� 1���� ��ñٷ��� 1�δ� 1����� �ٷνð�") Then 
	    '����������� ������ ��1������ñٷ���1�δ� 1����ձٷνð� 
		'- 24�ð� �ʰ��ϸ� ���� 
		'- �Ҽ��� 2�ڸ� �̸� ���� 
		if unicdbl(lgcTB_JT11_5.GetData("W7"),0) > 24 then 
		 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W7"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "����������� �������� 1���� ��ñٷ��� 1�δ� 1����� �ٷνð�", "24"))
		     blnError = True	
		end if
	Else
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W7"), 5, 2)

	If  ChkNotNull(lgcTB_JT11_5.GetData("W8"), "������������������� 1���� ��ñٷ��� 1�δ� 1����ձٷνð�") Then 
	    '������������������� 1���� ��ñٷ��� 1�δ� 1����ձٷνð� 
		'- 24�ð� �ʰ��ϸ� ���� 
		'- �Ҽ��� 2�ڸ� �̸� ���� 
		if unicdbl(lgcTB_JT11_5.GetData("W8"),0) > 24 then 
		 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W8"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "������������������� 1���� ��ñٷ��� 1�δ� 1����ձٷνð�", "24"))
		     blnError = True	
		end if
	Else
		blnError = True		
	End if	
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W8"), 5, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W9"), "��������������ñٷ��ڼ�") Then 
	    If unicdbl(lgcTB_JT11_5.GetData("W9"),0 ) <> unicdbl(lgcTB_JT11_5.GetData("W4"),0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W9"), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"��������������ñٷ��ڼ�", "�׸�(9)��������������ñٷ��ڼ�"))
	       blnError = True		
	    End if   
	Else
	 
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W9"), 7, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W10"), "���װ����� ��") Then 
	    ' �׸�(6)������뼼�װ����� + �׸�(10)����������װ����� 
		'- ���װ�����û��(A165)�� �ڵ�(91) ������� Ư�����װ��� �׸�(11)��󼼾װ� ��ġ 
		'(���װ����� �谡 ��0������ ū ��� �ݵ�� �Է�)
		
		Set cDataExists.A165  = new C_TB_JT1	' -- W6103MA1_HTF.asp �� ���ǵ� 
		Call SubMakeSQLStatements_W6113MA1("A165",iKey1, iKey2, iKey3)   
		cDataExists.A165.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A165.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 

			 If Not cDataExists.A165.LoadData then
			    blnError = True
				Call SaveHTFError(lgsPGM_ID, "��Ư �� 1ȣ  ���װ�����û��(A165)", TYPE_DATA_NOT_FOUND)	
			else
	
			  	if UNICDBL(lgcTB_JT11_5.GetData("W10"),  0) <> UNICDbl(cDataExists.A165.GetData("W5"), 0) then

			  	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������(13) "," ���װ�����û��(A165)�� �ڵ�(91) ������� Ư�����װ��� �׸�(11)��󼼾�"))
					blnError = True
					
			  	end if
			  	
			  	
			End if    
			Set cDataExists.A165 = Nothing	 
		

	Else
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W10"), 15, 0)
		

	sHTFBody = sHTFBody & UNIChar("", 10)	' -- ���� 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_JT11_5 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	   Case "A165" '-- �ܺ� ���� SQL	
	      ' ���װ�����û��(A165)�� �ڵ�(75) ����Ǿ����������������� ���װ��� �׸�(11) ��󼼾װ� ��ġ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " and A.W3 = '91'" & vbCrLf

	
	End Select
				Response.Write lgStrSQL
	PrintLog "SubMakeSQLStatements_W6113MA1 : " & lgStrSQL
End Sub

%>
