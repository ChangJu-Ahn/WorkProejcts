
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��16ȣ ���Աݾ� �������� 
'*  3. Program ID           : W2105MA1
'*  4. Program Name         : W2105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_16

Set lgcTB_16 = Nothing ' -- �ʱ�ȭ 

Class C_TB_16
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
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
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("D1",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("D2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
				
		If blnData1 = False And blnData2 = False And blnData3 = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function
	
	
	Function Clone(Byref pRs)
	   Set pRs = lgoRs1.clone   '���� 
    End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
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
	
	Function MoveFist(Byval pType)
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
		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
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
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_16H	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_16D1	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				
		  Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_16D2	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W2105MA1
	Dim A111

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W2105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
  ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W2105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W2105MA1"

	Set lgcTB_16 = New C_TB_16		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_16.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W2105MA1

	'==========================================
	' -- ��16ȣ ���Աݾ� �������� �������� 
	' -- 1. ����������� 
	iSeqNo = 1	
	
	Do Until lgcTB_16.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "SEQ_NO"), 6)
		End If
				
		If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(1, "W1_NM"), "���� �׸�") Then blnError = True	
			If Not ChkNotNull(lgcTB_16.GetData(1, "W2_NM"), "���� �׸�") Then blnError = True	
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "W1_NM"), 50)
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "W2_NM"), 50)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W3"), "��꼭�� ���Աݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W3"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W4"), "��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W5"), "��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W5"), 15, 0)
		
		If  ChkNotNull(lgcTB_16.GetData(1, "W6"), "������ ���Աݾ�") Then
		
			If unicdbl(lgcTB_16.GetData(1, "W6"),0) <> unicdbl(lgcTB_16.GetData(1, "W3"),0) + unicdbl(lgcTB_16.GetData(1, "W4"),0) - unicdbl(lgcTB_16.GetData(1, "W5"),0)  then
			    Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_16.GetData(1, "W6"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������ ���Աݾ�","�׸�(3)��꼭����Աݾ� + �׸�(4)����_���� - �׸�(5)����_����"))
				 blnError = True	
			End if
			
			If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) = 999999 Then
			   Set cDataExists.A111  = new C_TB_17		' -- W2107MA1_HTF.asp �� ���ǵ� 
			 '  Call SubMakeSQLStatements_W2105MA1("A111",iKey1, iKey2, iKey3)   
			   cDataExists.A111.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			   cDataExists.A111.WHERE_SQL = ""	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			   
					If Not cDataExists.A111.LoadData() Then
					       blnError = True
						  Call SaveHTFError(lgsPGM_ID, "��17ȣ ������ ���Աݾ� ���� ", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
					Else
				
					    
				        Call cDataExists.A111.Find ("1" ,"Code_no = '99'")   '�ڵ尡 99�ΰͰ� �� 
					   If unicdbl(cDataExists.A111.GetData(1,"W4"),0) <> unicdbl(lgcTB_16.GetData(1, "W6"),0) then
					      Call SaveHTFError(lgsPGM_ID,unicdbl(cDataExists.A111.getdata(1,"W4"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������ ���Աݾ�","��17ȣ ������ ���Աݾ� ������ �ڵ� 99�� �ݾ�"))
					       blnError = True
					   End if
					   
					End if
				Set  cDataExists.A111 = nothing
			   
			End If
		Else
		    blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W6"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "DESC1"), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- ���� 
		
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop

	

	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	' -- 2. �۾�������� ���Ѽ��Աݾ� 
	iSeqNo = 1	: blnError = False
sHTFBody = ""
	Do Until lgcTB_16.EOF(2) 
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "SEQ_NO"), 40)
		End If
		
		
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then		 
		   If Not ChkNotNull(lgcTB_16.GetData(2, "W7"), "�����") Then blnError = True	
		End if
		   
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "W7"), 50)
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(2, "W8"), "������") Then blnError = True	
		End if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "W8"), 30)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W9"), "���ޱݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W9"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W10"), "���ػ�������� �Ѱ��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W11"), "�Ѱ��翹����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W11"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W12"), "�����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W12"), 5, 2)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W13"), "�ͱݻ��Ծ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W13"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W14"), "���⸻���԰���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W14"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W15"), "���ȸ����԰���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W15"), 15, 0)
		
		If  ChkNotNull(lgcTB_16.GetData(2, "W16"), "������") Then 
		   '�׸�(13)�ͱݻ��Ծ� - �׸�(14)���⸻���԰��� - �׸�(15)���ȸ����԰��� +10,000 ~ -10,000
		   sTmp= unicdbl(lgcTB_16.GetData(2, "W13"),0) + unicdbl(lgcTB_16.GetData(2, "W14"),0) -unicdbl(lgcTB_16.GetData(2, "W15"),0) 
		   If unicdbl(lgcTB_16.GetData(2, "W16"),0) = sTmp + 10000 and unicdbl(lgcTB_16.GetData(2, "W16"),0) >= sTmp - 10000 then
		   Else
		       Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_16.GetData(2, "W16"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������","�׸�(13)�ͱݻ��Ծ� - �׸�(14)���⸻���԰��� - �׸�(15)���ȸ����԰��� +10,000 ~ -10,000"))
				blnError = True
		   End if
		   
		Else
		   blnError = True	
		End if
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W16"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 48) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(2)	' -- 1�� ���ڵ�� 
	Loop
			

	PrintLog "Write2File : " & sHTFBody
 
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If


		' -- 3. ��Ÿ ���Ա� 
	iSeqNo = 1	: blnError = False
	sHTFBody = ""
	Do Until lgcTB_16.EOF(3) 
		sHTFBody = sHTFBody & "85"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_16.GetData(3, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "SEQ_NO"), 40)
		End If
		'Response.End 
		
		'zzzz �հ�� üũ �ʿ����?
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "W17"), "����") Then blnError = True
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "W17"), 50)
		
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "W18"), "�ٰŹ���") Then blnError = True	
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "W18"), 80)
		
		If Not ChkNotNull(lgcTB_16.GetData(3, "W19"), "���Աݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(3, "W19"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(3, "W20"), "��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(3, "W20"), 15, 0)
		
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "DESC2"), "���") Then blnError = True	
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "DESC2"), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(3)	' -- 1�� ���ڵ�� 
	Loop

	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
			
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_16 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W2105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A111" '-- �ܺ� ���� SQL
           lgStrSQL = ""
      
	End Select
	PrintLog "SubMakeSQLStatements_W2105MA1 : " & lgStrSQL
End Sub
%>
