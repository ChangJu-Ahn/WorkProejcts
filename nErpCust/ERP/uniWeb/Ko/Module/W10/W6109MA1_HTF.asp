<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��Ư��2ȣ2 ����Ǿ��������������װ��� 
'*  3. Program ID           : W6109MA1
'*  4. Program Name         : W6109MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_JT2_2

Set lgcTB_JT2_2 = Nothing	' -- �ʱ�ȭ 

Class C_TB_JT2_2
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

	            If WHERE_SQL = "" Then	' �ܺ�ȣ���� �ƴϸ� ���������� ������ �ҷ��´� 
					lgStrSQL = lgStrSQL & " , B.IND_TYPE " & vbCrLf
	            End If
	            	            
				lgStrSQL = lgStrSQL & " FROM TB_JT2_2_200603	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 

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
Class TYPE_DATA_EXIST_W6109MA1
	Dim A165

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6109MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6109MA1"
	
	Set lgcTB_JT2_2 = New C_TB_JT2_2		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_JT2_2.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
	
	Set cDataExists = new  TYPE_DATA_EXIST_W6109MA1	
	
	'==========================================
	' -- ��Ư��2ȣ2 ����Ǿ��������������װ��� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("IND_TYPE"), "��������_����") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT2_2.GetData("IND_TYPE"), 50)

' -- ���İ������� ������ : 200603 
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1"), "ȯ����.�ǸŴ���߽��Ƿڼ� �����ݾ׹� �����������ī�� ���ݾ׹� �ܻ����ä�Ǵ㺸�������� �̿�ݾ�") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1"), 15, 0)
'	if  UNICDbl(lgcTB_JT2_2.GetData("W1"),  0) <> UNICDbl(lgcTB_JT2_2.GetData("W1_A"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_B"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_C"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_D"),0) then
'	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ȯ����.�ǸŴ���߽��Ƿڼ� �����ݾ׹� �����������ī�� ���ݾ׹� �ܻ����ä�Ǵ㺸�������� �̿�ݾ�","ȯ���������ݾ� + �ǸŴ���߽��Ƿڼ������ݾ� + �����������ī����ݾ� +�ܻ����ä�Ǵ㺸���������̿��"))
'	    blnError = True
'	end if
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_A"), "ȯ���������ݾ�") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_A"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_B"), "�ǸŴ���߽��Ƿڼ������ݾ�") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_B"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_C"), "�����������ī����ݾ�") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_C"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_D"), "�ܻ����ä�Ǵ㺸���������̿�ݾ�") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_D"), 15, 0)

'	'�����ݾ� : (�׸�(8)ȯ�����ǸŴ���߽��Ƿڼ������ݾ�,�����������ī����ݾ� �׿ܻ����ä�Ǵ㺸���� �̿�ݾ� - �׸�(9)��Ӿ��������ݾ�) x 3 /1000 
'	If  ChkNotNull(lgcTB_JT2_2.GetData("W3"), "�����ݾ�") Then 
'	    if  UNICDbl(lgcTB_JT2_2.GetData("W3"),  0)  <> Fix((UNICDbl(lgcTB_JT2_2.GetData("W1"), 0) - UNICDbl(lgcTB_JT2_2.GetData("W2"),0))* UNICDbl(lgcTB_JT2_2.GetData("W3_RATE_VALUE"),0))  then
'	         Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����ݾ�","  (�׸�(8)ȯ�����ǸŴ���߽��Ƿڼ������ݾ�,�����������ī����ݾ� �׿ܻ����ä�Ǵ㺸���� �̿�ݾ� - �׸�(9)��Ӿ��������ݾ�) x 3 /1000(�Ҽ�������) "))
'	         blnError = True
'	    End if
'	Else	
'		blnError = True		
'	End if	
'--------------------------------------
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W11"), "��Ӿ��������ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W11"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_HAP_C"), "�����ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_HAP_C"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W13"), "���⼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W13"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W14"), "�ѵ���") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W14"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W15"), "��������") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W15"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_SUM"), "���ݾ�_�հ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_SUM"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_1"), "���ݾ�_���ޱ���1�հ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_2"), "���ݾ�_���ޱ���2�հ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_2"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_GA_C"), "�������ݾ�_��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_GA_C"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_NA_C"), "�������ݾ�_��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_NA_C"), 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 44)	' -- ���� 


	' -- ���� : �����ݾ� = a * b �� Sum
	if  UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W12_GA_C"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W12_NA_C"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W12_GA_C"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(12) �����ݾ�"," �����ݾ�( (a) X (b) )�� �� + ��"))
	     blnError = True
	end if
	
	' -- ���� : (15)���������� (12)�� (14)�� ���� �ݾ� 
	If UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) < UNICDbl(lgcTB_JT2_2.GetData("W14"), 0) Then
		if  UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) <> UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15) ��������","(12)��������"))
		     blnError = True
		end if
	Else
		if  UNICDbl(lgcTB_JT2_2.GetData("W14"), 0) <> UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15) ��������","(14)�ѵ���"))
		     blnError = True
		end if
	End If
	
	
	' -- ���� : W15 �ݾװ� ��ġ���� 
	If  UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) > 0 Then 
	
		Set cDataExists.A165  = new C_TB_JT1	' -- W6103MA1_HTF.asp �� ���ǵ� 
		
		Call SubMakeSQLStatements_W6109MA1("A165",iKey1, iKey2, iKey3)   
		
		cDataExists.A165.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A165.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
	
	   ' ��������(13) : �׸�(10)�����ݾװ� �׸�(12)�ѵ����߿� �����ݾ��� �������׿� �Է��մϴ�.
       ' ���װ�����û��(A165)�� �ڵ�(75) ����Ǿ����������������� ���װ��� �׸�(11) ��󼼾�  �� ��ġ 

        If Not cDataExists.A165.LoadData then
            blnError = True
			Call SaveHTFError(lgsPGM_ID, "��Ư �� 1ȣ  ���װ�����û��(A165)", TYPE_DATA_NOT_FOUND)	
        else

		  	if UNICDBL(lgcTB_JT2_2.GetData("W15"),  0) <> UNICDbl(cDataExists.A165.GetData("W5"), 0) then

		  	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15)�������� "," ���װ�����û��(A165)�� �ڵ�(75) ����Ǿ����������������� ���װ��� �׸�(6) ��������"))
				blnError = True
				
		  	end if
		  	
		  	
		End if   
		Set cDataExists.A165 = Nothing	   
	Else
	    blnError = True
	End if    		



	
	' -- ����� �������������� ���� �������װ�꼭 - ���ݾ� 
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "1"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_SUM"), "ȯ���� �����ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_1"), "ȯ���� �����ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_2"), "ȯ���� �����ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 
	
	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_A_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_A_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_A_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_A_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ȯ���� �����ݾ� �հ�","ȯ���� �����ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if
	
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "2"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_SUM"), "�ǸŴ���߽��Ƿڼ� �����ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_1"), "�ǸŴ���߽��Ƿڼ� �����ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_2"), "�ǸŴ���߽��Ƿڼ� �����ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 

	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_B_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_B_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_B_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_B_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ǸŴ���߽��Ƿڼ� �����ݾ� �հ�","�ǸŴ���߽��Ƿڼ� �����ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "3"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_SUM"), "�����������ī�� ���ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_1"), "�����������ī�� ���ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_2"), "�����������ī�� ���ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 

	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_C_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_C_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_C_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_C_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����������ī�� ���ݾ� �հ�","�����������ī�� ���ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "4"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_SUM"), "�ܻ����ä�Ǵ㺸�������� �̿�ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_1"), "�ܻ����ä�Ǵ㺸�������� �̿�ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_2"), "�ܻ����ä�Ǵ㺸�������� �̿�ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 

	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_D_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_D_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_D_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_E_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ܻ����ä�Ǵ㺸�������� �̿�ݾ� �հ�","�ܻ����ä�Ǵ㺸�������� �̿�ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "5"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_SUM"), "���ŷ����� �̿�ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_1"), "���ŷ����� �̿�ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_2"), "���ŷ����� �̿�ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 

	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_E_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_E_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_E_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_E_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���ŷ����� �̿�ݾ� �հ�","���ŷ����� �̿�ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	sHTFBody = sHTFBody & "6"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_SUM"), "��Ʈ��ũ������ �̿�ݾ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_1"), "��Ʈ��ũ������ �̿�ݾ� 30�� �̳�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_2"), "��Ʈ��ũ������ �̿�ݾ� 31�� ~ 60��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- ���� 


	' -- ���� : ���ݾ� �հ� = 30�� �̳� + 31�� ~ 60�� 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_F_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_F_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_F_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_F_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��Ʈ��ũ������ �̿�ݾ� �հ�","��Ʈ��ũ������ �̿�ݾ��� ���ޱ��� 30���̳� + 31��~60��"))
	     blnError = True
	end if

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_JT2_2 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	 Case "A165" '-- �ܺ� ���� SQL	
	      ' ���װ�����û��(A165)�� �ڵ�(75) ����Ǿ����������������� ���װ��� �׸�(11) ��󼼾װ� ��ġ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " and A.W3 = '75'" & vbCrLf
	
	End Select
	PrintLog "SubMakeSQLStatements_W6109MA1 : " & lgStrSQL
End Sub

%>
