<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��Ư��3ȣ�������η°��߸��� 
'*  3. Program ID           : W6111MA1
'*  4. Program Name         : W6111MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_JT3

Set lgcTB_JT3 = Nothing ' -- �ʱ�ȭ 

Class C_TB_JT3
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
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
		
		If blnData1 = False And blnData2 = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
	On Error Resume Next   

		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
					
				End If
		End Select
		
		if err.number <> 0 then
					   Response.Write pFieldNm
					   Response.End 
					End if 
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
		Call SubCloseRs(lgoRs2)		
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_JT3A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "B"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_JT3B	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6111MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6111MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6111MA1"

	Set lgcTB_JT3 = New C_TB_JT3		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_JT3.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W6111MA1

	'==========================================
	' -- ��Ư��3ȣ�������η°��߸��� �������� 
	' -- 1. ����׸��԰ŷ��� 
	'==========================================
	iSeqNo = 1	
	'Response.End 'zzz
	Do Until lgcTB_JT3.EOF(2) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_JT3.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "ACCT_NM"), "��������") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W1_T"), "���йװ���(6)") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W2_T"), "���йװ���(7)") Then blnError = True
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W3_T"), "���йװ���(8)") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W4_T"), "���йװ���(9)") Then blnError = True	
			 
			 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "SEQ_NO"), 6)
			
			  
		End If
					
		 If UNICDbl(lgcTB_JT3.GetData(2, "W6"),0) <> UNICDbl(lgcTB_JT3.GetData(2, "W1"),0) + UNICDbl(lgcTB_JT3.GetData(2, "W2"),0) + UNICDbl(lgcTB_JT3.GetData(2, "W3"),0)+ UNICDbl(lgcTB_JT3.GetData(2, "W4"),0)+ UNICDbl(lgcTB_JT3.GetData(2, "W5"),0) Then
		   '���й׺���� �׸�(6)�ݾ� + (7) + (8) + (9) + (10)�� �ݾ��� �� 
		    Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_JT3.GetData(2, "W6"),0) ,UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����" & lgcTB_JT3.GetData(2, "ACCT_NM")  & "�հ�","�� ���� ����� ��"))
			blnError = True	
		 End If
	

		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "ACCT_NM"), 20)
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W1_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W1"), "�ݾ�(6)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W1"), 15, 0)

			
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W2_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W2"), "�ݾ�(7)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W2"), 15, 0)

		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W3_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W3"), "�ݾ�(8)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W3"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W4_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W4"), "�ݾ�(9)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W4"), 15, 0)

		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W5"), "�ݾ�(10)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W5"), 15, 0)

		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W6"), "�ݾ�_") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W6"), 15, 0)

		sHTFBody = sHTFBody & UNIChar("", 48) & vbCrLf	' -- ���� 

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_JT3.MoveNext(2)	' -- 2�� ���ڵ�� 
	Loop

	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""

	'==========================================
	' 2 �������η°��ߺ����_�������η°��ߺ��������߻����ǰ��,�������� 
	'==========================================
	
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D1_S"), "����4�Ⱓ�߻��հ��_�Ⱓ1���۳����") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D1_S"), "����4�Ⱓ�߻��հ��_�Ⱓ1���۳����") Then
	       blnError = True	
	    End if
	   
	Else
	    blnError = True	
	End If    

	  	 sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D1_S"))
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D1_E"), "����4�Ⱓ�߻��հ��_�Ⱓ1��������") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D1_E"), "����4�Ⱓ�߻��հ��_�Ⱓ1��������") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D1_E"))
			  
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT1"), "����4�Ⱓ�߻��հ��_�Ⱓ1�ݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT1"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D2_S"), "����4�Ⱓ�߻��հ��_�Ⱓ2���۳����") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D2_S"), "����4�Ⱓ�߻��հ��_�Ⱓ2���۳����") Then
	       blnError = True	
	    End if
	Else
		blnError = True	
	End if	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D2_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D2_E"), "����4�Ⱓ�߻��հ��_�Ⱓ2��������") Then 
	   If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D2_E"), "����4�Ⱓ�߻��հ��_�Ⱓ2��������") Then
	       blnError = True	
	    End if
	   
	Else
	    blnError = True	
	End if   
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D2_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT2"), "����4�Ⱓ�߻��հ��_�Ⱓ2�ݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT2"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D3_S"), "����4�Ⱓ�߻��հ��_�Ⱓ3���۳����") Then
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D3_S"), "����4�Ⱓ�߻��հ��_�Ⱓ3���۳����") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D3_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D3_E"), "����4�Ⱓ�߻��հ��_�Ⱓ3��������") Then
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D3_E"), "����4�Ⱓ�߻��հ��_�Ⱓ3��������") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D3_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT3"), "����4�Ⱓ�߻��հ��_�Ⱓ3�ݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT3"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D4_S"), "����4�Ⱓ�߻��հ��_�Ⱓ4���۳����") Then
	     If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D4_S"), "����4�Ⱓ�߻��հ��_�Ⱓ4���۳����") Then
	       blnError = True	
	    End if
	    
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D4_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D4_E"), "����4�Ⱓ�߻��հ��_�Ⱓ4��������") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D4_E"), "����4�Ⱓ�߻��հ��_�Ⱓ4��������") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D4_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT4"), "����4�Ⱓ�߻��հ��_�Ⱓ4�ݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT4"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_SUM"), "����4�Ⱓ�߻��հ��_��") Then 
	   If UNICDbl(lgcTB_JT3.GetData(1, "W8_SUM"),0) <> UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT1"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT2"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT3"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT4"),0) Then
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W8_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����4�Ⱓ�߻��հ��_��","�� �Ⱓ���� �ݾ��� ��"))
			blnError = True	
	   End If
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_SUM"), 15, 0)

	'�׸�(14)����4�Ⱓ ����չ߻��� = �׸�(13)����4�Ⱓ�߻��հ��_�� X (48/����4�Ⱓ�ǻ����������) X (1/4)X (���ؿ�������/12)
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W9"), "����4�Ⱓ����չ߻�") Then
	    If Unicdbl(lgcTB_JT3.GetData(1, "W9"),0) <> fix(Unicdbl(lgcTB_JT3.GetData(1, "W8_SUM"),0) * (48/Unicdbl(lgcTB_JT3.GetData(1, "W_4Year_Mth"),0))*0.25* (Unicdbl(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"),0)/12) ) Then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W9"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����4�Ⱓ ����չ߻���","����4�Ⱓ�߻��հ��_�� X (48/����4�Ⱓ�ǻ����������) X (1/4)X (���ؿ�������/12)"))
	    	blnError = True	
	    End if
	Else
	   blnError = True	
	End if   
	 

	sHTFBody = sHTFBody &  UNINumeric(lgcTB_JT3.GetData(1, "W9"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), "����4�Ⱓ�ǻ����������") Then 
	    If UNICDbl(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"),0) > 48 then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "����4�Ⱓ�ǻ����������","48"))
	       blnError = True	
	    End if
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), 2, 0)
	
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), "���ؿ�������") Then 
	   If UNICDbl(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"),0) > 12 then
	  
	   
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "���ؿ�������","12"))
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), 2, 0)

    '�׸�(15)�����߻��ݾ�	= �׸�(12)���ؿ����ǿ������η°��ߺ�߻����� �ݾ�(��) �հ� -�׸�(14)����4�Ⱓ����չ߻��� 
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W10"), "�����߻��ݾ�") Then
	   ' IF UNICDbl(lgcTB_JT3.GetData(1, "W10"),0) <>  UNICDbl(lgcTB_JT3.GetData(1, "W15_11"),0) - Unicdbl(lgcTB_JT3.GetData(1, "W9"),0) Then
	    '   Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����߻��ݾ�","���ؿ����ǿ������η°��ߺ�߻����� �ݾ�(��) �հ� - ����4�Ⱓ����չ߻���"))
	     '  blnError = True	
	   ' End if

	Else
	     blnError = True	
	End If     
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W10"), 15, 0)
	'	
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_11"), "���ؿ����ѹ߻��ݾװ���_���ݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W15_11"), 15, 0)

	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_12_VALUE"), "���ؿ����ѹ߻��ݾװ���_������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(UNICDbl(lgcTB_JT3.GetData(1, "W15_12_VALUE"),0) * 100, 5, 2)
		
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_13"), "���ؿ����ѹ߻��ݾװ���_��������") Then blnError = True	
	
	'�׸�(20)��(18)���ؿ����ѹ߻��ݾװ���_��������	=�׸�(20)�� (16)���ؿ����ѹ߻��ݾװ���_���ݾ� X �׸�(20)�� (17)���ؿ����ѹ߻��ݾװ���_������(15/100)(��������10000)	
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W15_13"), "���ؿ����ѹ߻��ݾװ���_��������") Then
	    sTmp =  UNICDbl(lgcTB_JT3.GetData(1, "W15_11"),0) * UNICDbl(lgcTB_JT3.GetData(1, "W15_12_Value"),0) 
	    If (sTmp - 10000) <=  Unicdbl(lgcTB_JT3.GetData(1, "W15_13"),0)  And  Unicdbl(lgcTB_JT3.GetData(1, "W15_13"),0)  <= (sTmp + 10000)  Then
	    Else
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W15_13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���ؿ����ѹ߻��ݾװ���_��������","���ؿ����ѹ߻��ݾװ���_���ݾ� * ���ؿ����ѹ߻��ݾװ���_������"))
	    End if
	Else
	    blnError = True	
	End If    
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W15_13"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W15_14"), 30)'"���ؿ����ѹ߻��ݾװ���_��� 
    
    '�׸�(21)��(16)�����߻��ݾװ���_���ݾ�	= �׸�(15)�����߻��ݾ� 
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W16_11"), "�����߻��ݾװ���_���ݾ�") Then 
	
	    If UNICDbl( lgcTB_JT3.GetData(1, "W16_11"),0) <> UNICDbl( lgcTB_JT3.GetData(1, "W10"),0) Then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W16_11"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����߻��ݾװ���_���ݾ�","�����߻��ݾ�"))
	       blnError = True
	    End if   
	   
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W16_11"), 15, 0)
		
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W16_12_Value"), "�����߻��ݾװ���_������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(UNICDbl(lgcTB_JT3.GetData(1, "W16_12_Value"),0) * 100, 5, 2)
	
	'�׸�(21)��(18)���ؿ����ѹ߻��ݾװ���_��������	- �׸�(21)�� (16)���ؿ����ѹ߻��ݾװ���_���ݾ� X �׸�(21)��   (17)�����߻��ݾװ���_������(40/100,�߼ұ���ǰ�� 50/100)(��������10000)


	If  ChkNotNull(lgcTB_JT3.GetData(1, "W16_13"), "�����߻��ݾװ���_��������") Then
	    sTmp =  UNICDbl(lgcTB_JT3.GetData(1, "W16_11"),0) * UNICDbl(lgcTB_JT3.GetData(1, "W16_12_Value"),0) 
	    If (sTmp - 10000) <=  Unicdbl(lgcTB_JT3.GetData(1, "W16_13"),0)  And  Unicdbl(lgcTB_JT3.GetData(1, "W16_13"),0)  <= (sTmp + 10000)  Then
	    Else
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W16_13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���ؿ����ѹ߻��ݾװ���_��������","���ؿ����ѹ߻��ݾװ���_���ݾ� * ���ؿ����ѹ߻��ݾװ���_������"))
	    End if
	Else
	    blnError = True	
	End If    
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W16_13"), 15, 0)


	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W16_14"), 30)   '�����߻��ݾװ���_��� 
  
    '�׸�(22)��(18)���ؿ�����������������_��������	- �߼ұ��(�׸�(20)�� (18)���ؿ����ѹ߻��ݾװ���_�������װ� �׸�(21)��   (18)�����߻��ݾװ���_���������� ����) 
    '�Ǵ� �߼ұ�������߼ұ��(�׸�(21)  �� (18) �����߻��ݾװ���_��������)  ���װ�����û��(A165)�� �ڵ�(32) ���� �� �η°��ߺ� ���װ��� �׸�(11) ��  �󼼾װ� ��ġ 
    ' (�׸�(22)��(18)���ؿ�����������������_���������� ��0������ ū ��� �ݵ�� �Է�)
	
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W17"), "���ؿ�����������������_��������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W17"), 15, 0)

	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W17_14"), 60) '���ؿ�����������������_��� 


	sHTFBody = sHTFBody & UNIChar("", 16) & vbCrLf	' -- ���� 

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_JT3 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL
			

	End Select
	PrintLog "SubMakeSQLStatements_W6111MA1 : " & lgStrSQL
End Sub
%>
