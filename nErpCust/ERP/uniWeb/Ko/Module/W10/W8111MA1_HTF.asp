<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��12ȣ ��Ư�� ����ǥ�ع� ����������꼭 
'*  3. Program ID           : W8111MA1
'*  4. Program Name         : W8111MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_12

Set lgcTB_12 = Nothing	' -- �ʱ�ȭ 

Class C_TB_12
	' -- ���̺��� �÷����� 
	Dim W5_AMT
	Dim W5_RATE
	Dim W5_RATE_VAL
	Dim W5_TAX
	Dim W6
	Dim W6_AMT
	Dim W6_RATE
	Dim W6_RATE_VAL
	Dim W6_TAX
	Dim W7
	Dim W7_AMT
	Dim W7_RATE
	Dim W7_RATE_VAL
	Dim W7_TAX
	Dim W8_AMT
	Dim W8_TAX
	Dim W10_AMT
	Dim W10_RATE
	Dim W10_RATE_VAL
	Dim W10_TAX
	Dim W11
	Dim W11_AMT
	Dim W11_RATE
	Dim W11_RATE_VAL
	Dim W11_TAX
	Dim W12_AMT
	Dim W12_TAX
			
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
		
		W5_AMT			= oRs1("W5_AMT")
		W5_RATE			= oRs1("W5_RATE")
		W5_RATE_VAL		= oRs1("W5_RATE_VAL")
		W5_TAX			= oRs1("W5_TAX")
		W6				= oRs1("W6")
		W6_AMT			= oRs1("W6_AMT")
		W6_RATE			= oRs1("W6_RATE")
		W6_RATE_VAL		= oRs1("W6_RATE_VAL")
		W6_TAX			= oRs1("W6_TAX")
		W7				= oRs1("W7")
		W7_AMT			= oRs1("W7_AMT")
		W7_RATE			= oRs1("W7_RATE")
		W7_RATE_VAL		= oRs1("W7_RATE_VAL")
		W7_TAX			= oRs1("W7_TAX")
		W8_AMT			= oRs1("W8_AMT")
		W8_TAX			= oRs1("W8_TAX")
		W10_AMT			= oRs1("W10_AMT")
		W10_RATE		= oRs1("W10_RATE")
		W10_RATE_VAL	= oRs1("W10_RATE_VAL")
		W10_TAX			= oRs1("W10_TAX")
		W11				= oRs1("W11")
		W11_AMT			= oRs1("W11_AMT")
		W11_RATE		= oRs1("W11_RATE")
		W11_RATE_VAL	= oRs1("W11_RATE_VAL")
		W11_TAX			= oRs1("W11_TAX")
		W12_AMT			= oRs1("W12_AMT")
		W12_TAX			= oRs1("W12_TAX")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_12	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W8111MA1
	Dim A151

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8111MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8111MA1"
	
	Set lgcTB_12 = New C_TB_12		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_12.LoadData	Then Exit Function		' -- ��12ȣ ���� �ε� 
	
	Set cDataExists = new TYPE_DATA_EXIST_W8111MA1
	'==========================================
	' -- ��12ȣ ��Ư�� ����ǥ�ع� ����������꼭 �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	If ChkNotNull(lgcTB_12.W5_AMT, "�Ϲݹ���_����ǥ�رݾ�")  Then ' -- ����Ÿ����� ������ 
	
			
			' --��13ȣ�����Ư����������󰨸鼼���հ�ǥ 
			Set cDataExists.A151  = new C_TB_13	' -- W8109MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8111MA1("A105",iKey1, iKey2, iKey3)   
			
			cDataExists.A151.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A151.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A151.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��13ȣ�����Ư����������󰨸鼼���հ�ǥ", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				cDataExists.A151.FIND 1, " W1_CD='10' "
				If  (UNICDbl(lgcTB_12.W5_AMT, 0) <> UNICDbl(cDataExists.A151.GetData(1, "W4"), 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_12.W5_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[�Ϲݹ���] ����ǥ�رݾ�","�����Ư����������󰨸鼼���հ�ǥ��(A151)�� �׸�(10)�� �ݾ�"))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A151 = Nothing
	
	Else
		blnError = True
	End If

	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_AMT, 15, 0)
	
	If Not ChkNotNull(lgcTB_12.W5_RATE, "�Ϲݹ���_����") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_RATE, 5, 2)

	
	If ChkNotNull(lgcTB_12.W5_TAX, "�Ϲݹ���_����") Then 
	   if  UNICDbl(lgcTB_12.W5_TAX, 0) <>   Fix((UNICDbl(lgcTB_12.W5_AMT, 0) * UNICDbl(lgcTB_12.W5_RATE_VAL,0))) then
	       blnError = True	
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W5_TAX & " <> " & Int((UNICDbl(lgcTB_12.W5_AMT, 0) * UNICDbl(lgcTB_12.W5_RATE_VAL,0))), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[�Ϲݹ���]����","�׸�(5)�� ����ǥ�رݾ� �� ����(" & lgcTB_12.W5_RATE &")"))
	   end if
	else

		blnError = True	
	end if	
	   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_TAX, 15, 0)
	
	'��Ÿ1���� 
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W6, 20)
	
	
	
	If  ChkNotNull(lgcTB_12.W6_AMT, "�Ϲݹ���(��Ÿ1)_" & lgcTB_12.W6 & "_����") Then 
	    if UNICDbl(lgcTB_12.W6_AMT, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ1_�ݾ� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_AMT, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W6_RATE, "�Ϲݹ���(��Ÿ1)_" & lgcTB_12.W6 & "_�ݾ�") Then 
	    if UNICDbl(lgcTB_12.W6_RATE, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ1_���� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_RATE, 5, 2)
	
	
	
	If  ChkNotNull(lgcTB_12.W6_TAX, "�Ϲݹ���(��Ÿ1)_" & lgcTB_12.W6 & "_����") Then 
	    if UNICDbl(lgcTB_12.W6_TAX, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ1_���� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_TAX, 15, 0)
	
	'��Ÿ2���� 
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W7, 20)
	
	
	
	If  ChkNotNull(lgcTB_12.W7_AMT, "�Ϲݹ���(��Ÿ2)_" & lgcTB_12.W7 & "_�ݾ�") Then 
	    if UNICDbl(lgcTB_12.W7_AMT, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ2_�ݾ� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_AMT, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W7_RATE, "�Ϲݹ���(��Ÿ2)_" & lgcTB_12.W7 & "_����") Then 
	    if UNICDbl(lgcTB_12.W7_RATE, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ2_���� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_RATE, 5, 2)
	
	
	
	If  ChkNotNull(lgcTB_12.W7_TAX, "�Ϲݹ���(��Ÿ1)_" & lgcTB_12.W7 & "_����") Then 
	    if UNICDbl(lgcTB_12.W7_TAX, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[�Ϲݹ���]��Ÿ1_�ݾ� ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_TAX, 15, 0)
	
	
	
	If ChkNotNull(lgcTB_12.W8_AMT, "�Ϲݹ���_�Ұ�_����ǥ�رݾ�") Then 
	   If UNICDbl(lgcTB_12.W8_AMT,  0) <> UNICDbl(lgcTB_12.W5_AMT, 0) + UNICDbl(lgcTB_12.W6_AMT,  0) + UNICDbl(lgcTB_12.W7_AMT,  0) then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[�Ϲݹ���]�Ұ�_����ǥ�رݾ�","�׸�(5)�� ����ǥ�رݾ� + �׸�(6)�� ����ǥ�رݾ� + �׸�(7)�� ����ǥ�رݾ�"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W8_AMT, 15, 0)
	
	If ChkNotNull(lgcTB_12.W8_TAX, "�Ϲݹ���_�Ұ�_����ǥ�رݾ�") Then 
	   If UNICDbl(lgcTB_12.W8_TAX, 0) <> UNICDbl(lgcTB_12.W5_TAX,0) + UNICDbl(lgcTB_12.W6_TAX, 0) + UNICDbl(lgcTB_12.W7_TAX, 0) then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[�Ϲݹ���]�Ұ�_����","�׸�(5)�� ���� + �׸�(6)�� ���� + �׸�(7)�� ����"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W8_TAX, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W10_AMT, "���չ��ε�_����ǥ�رݾ�") Then 
	    	' --��13ȣ�����Ư����������󰨸鼼���հ�ǥ 
			Set cDataExists.A151  = new C_TB_13	' -- W8109MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8111MA1("A151",iKey1, iKey2, iKey3)   
			
			cDataExists.A151.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A151.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A151.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��13ȣ�����Ư����������󰨸鼼���հ�ǥ", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				    
				  
				If  (UNICDbl(lgcTB_12.W10_AMT, 0) <> UNICDbl(cDataExists.A151.GetData(2, "W7"), 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_12.W10_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[���չ��ε�]����ǥ�رݾ�","�����Ư����������󰨸鼼���հ�ǥ(A151)�� ���չ��ε��� ���鼼���׸�(7)�� ��"))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A151 = Nothing
	else
	
	   blnError = True	
	end if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_AMT, 15, 0)
	
	
	
	If Not ChkNotNull(lgcTB_12.W10_RATE, "���չ��ε�_����") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_RATE, 5, 2)
	
	If ChkNotNull(lgcTB_12.W10_TAX, "���չ��ε�_����") Then 
	   if  UNICDbl(lgcTB_12.W10_TAX,  0) <>  Fix(UNICDbl(lgcTB_12.W10_AMT,0) *  UNICDbl(lgcTB_12.W10_RATE, 2)) then
	       blnError = True	
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W10_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[���չ��ε�]����","�׸�(10)�� ����ǥ�رݾ� �� ����(20%) "))
	   end if
	else

		blnError = True	
	end if	
	   	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_TAX, 15, 0)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W11, 20)
	

	

    If  ChkNotNull(lgcTB_12.W11_AMT, "���չ��ε�(��Ÿ1)_�ݾ�")  then
	      if UNICDbl(lgcTB_12.W11_AMT,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[���չ��ε�]��Ÿ�ݾ׼��� ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_AMT, 15, 0)
	
	
	If  ChkNotNull(lgcTB_12.W11_RATE, "���չ��ε�(��Ÿ1)_����")  then
	    if UNICDbl(lgcTB_12.W11_RATE,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[���չ��ε�]��Ÿ1_���� ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_RATE, 5, 2)
	

	
	If  ChkNotNull(lgcTB_12.W11_TAX, "���չ��ε�(��Ÿ1)_����")  then
	    if UNICDbl(lgcTB_12.W11_TAX,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[���չ��ε�]��Ÿ1_���� ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_TAX, 15, 0)
	

	If ChkNotNull(lgcTB_12.W12_AMT, "���չ��ε�_�Ұ�_����ǥ�رݾ�") Then 
	   If UNICDbl(lgcTB_12.W12_AMT,  0) <> UNICDbl(lgcTB_12.W10_AMT, 0) + UNICDbl(lgcTB_12.W11_AMT, 0)  then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[���չ��ε�]�Ұ�_����ǥ�رݾ�","�׸�(10)�� ����ǥ�رݾ� + �׸�(11)�� ����ǥ�رݾ�"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W12_AMT, 15, 0)
	
	
	
	
	If Not ChkNotNull(lgcTB_12.W12_TAX, "���չ��ε�__����") Then blnError = True	
	If ChkNotNull(lgcTB_12.W12_TAX, "���չ��ε�_�Ұ�_����") Then 
	   If UNICDbl(lgcTB_12.W12_TAX, 0) <> UNICDbl(lgcTB_12.W10_TAX, 0) + UNICDbl(lgcTB_12.W11_TAX,  0)  then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W12_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[���չ��ε�]�Ұ�_����","�׸�(10)�� ���� + �׸�(11)�� ����"))
	   End if
	Else
	   
	   blnError = True	
	End if  
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W12_TAX, 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 49) 
	
	
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If
	
	'Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_12 = Nothing	' -- �޸�����  <-- W8101MA1_HTF���� ����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 

	  Case "A151" '-- �ܺ� ���� �ݾ� 
	
			lgStrSQL = ""
			'lgStrSQL = lgStrSQL & "	AND  A.W1_CD	= '10' 	" & vbCrLf
	
			
	End Select
	PrintLog "SubMakeSQLStatements_W8111MA1 : " & lgStrSQL
End Sub

%>
