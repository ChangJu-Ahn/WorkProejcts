<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��47ȣ �ֿ��������(��)
'*  3. Program ID           : W9103MA1
'*  4. Program Name         : W9103MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_47B

Set lgcTB_47B = Nothing ' -- �ʱ�ȭ 

Class C_TB_47B
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		
	Private lgoRs3		
	Private lgoRs4		
	
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3, blnData4
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' --������ �о�´�.
		Call SubMakeSQLStatements("1",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
			Exit Function
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
			
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("3",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("4",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs4,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData4 = False
		End If
		
		If blnData1 = False And blnData2 = False And blnData3 = False And blnData4 = False Then
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
			Case 3
				lgoRs3.Find pWhereSQL
			Case 4
				lgoRs4.Find pWhereSQL
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
			Case 4
				lgoRs4.Filter = pWhereSQL
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
			Case 4
				EOF = lgoRs4.EOF
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
			Case 4
				lgoRs4.MoveFirst
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
			Case 4
				lgoRs4.MoveNext
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
			Case 4
				If Not lgoRs4.EOF Then
					GetData = lgoRs4(pFieldNm)
				End If
		End Select
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
		Call SubCloseRs(lgoRs4)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
		Call SubCloseRs(lgoRs4)	
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT  "
			lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
			lgStrSQL = lgStrSQL & " FROM TB_47B" & pMode & "	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					
			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		
	 

			PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9103MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9103MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists ,sTmp1,Stmp2
    Dim iSeqNo, sMsg
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9103MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9103MA1"

	Set lgcTB_47B = New C_TB_47B		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_47B.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9103MA1

	'==========================================
	' -- ��47ȣ �ֿ��������(��) �������� 
	' -- 1. ����׸��԰ŷ��� 
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
	Do Until lgcTB_47B.EOF(1) 
	
		sTmp = lgcTB_47B.GetData(1, "W1")
		
		Select Case sTmp
			Case "101"
				sMsg = "��ǰ�׻�ǰ"
			Case "102"
				sMsg = "����ǰ�����ǰ"
			Case "103"
				sMsg = "�����"
			Case "104"
				sMsg = "����ǰ"
			Case "105"
				sMsg = "��������_ä��"
			Case "106"
				sMsg = "��������_��Ÿ"
			Case "107"
				sMsg = "�հ�"
		End Select
		
		
		
		
		
		 If sMsg = "101" Or  sMsg = "102"   Or  sMsg = "103" Or  sMsg = "104" Or  sMsg = "105" Then
		     sTmp1 ="1,2,3,4,5,6,7,8"
		  	  	'1:������, 2:���Լ����, 3:���Լ����, 4:����չ�, 5:�̵���չ�, 6:���Ⱑ��ȯ���� 7:������, 8:��Ÿ 
		 Else	
		    '1:������, 2:����չ�, 3:�̵���չ�, 4:�ð���,5:��Ÿ 
		    sTmp1 ="1,2,3,4,5,"
		 
		 End If
			   
			   
	     IF sTmp <> "107"  AND  UNICDbl(lgcTB_47B.GetData(1, "W4"),0)  <> 0 Then 
		   
		     If  ChkNotNull(lgcTB_47B.GetData(1, "W2"), sMsg & "_�Ű���") Then
		       
		         IF Not ChkBoundary(sTmp1, lgcTB_47B.GetData(1, "W2"),sMsg & "_�Ű���") Then
		 	       blnError = True	
		 	       
		 	    End if
		 	Else
		 	    blnError = True	
		 	End IF   
		       
		    
		   
		   If  ChkNotNull(lgcTB_47B.GetData(1, "W3"), sMsg & "_�򰡹��") Then 
				IF Not ChkBoundary(sTmp1, lgcTB_47B.GetData(1, "W3"),sMsg & "_�򰡹��") Then
				       blnError = True	
				
				End IF   
		   End if  
		   
		  
	  
	
		End IF   
	
		IF sTmp <> "107" Then
		    sHTFBody = sHTFBody & UNIChar(lgcTB_47B.GetData(1, "W2"), 1) '�Ű��� 
				
		   sHTFBody = sHTFBody & UNIChar(lgcTB_47B.GetData(1, "W3"), 1) '�򰡹�� 
		End If 
		

		If Not ChkNotNull(lgcTB_47B.GetData(1, "W4"), sMsg & "_ȸ����ݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W4"), 15, 0)
		
	
		If Not ChkNotNull(lgcTB_47B.GetData(1, "W5"), sMsg & "_�������ݾ�_�Ű���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W5"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(1, "W6"),  sMsg & "_�������ݾ�_���Լ����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W6"), 15, 0)
		
		'-������	- �׸�(5)�������ݾ�_�Ű��� - �׸�(4)ȸ����ݾ׶Ǵ� (�׸�(5)�� �׸�(6)�������ݾ�_���Լ���� �� ū �ݾ�-�׸�(4)ȸ����ݾ�)
		If  ChkNotNull(lgcTB_47B.GetData(1, "W7"), sMsg & "_������") Then 
		
		     If UniCDBL(lgcTB_47B.GetData(1, "W5"),0) > UNICDbl(lgcTB_47B.GetData(1, "W6"),0)  Then
		        Stmp2 = UNICDBl(lgcTB_47B.GetData(1, "W5"),0)  - UNICDbl(lgcTB_47B.GetData(1, "W4"),0)
		     Else
		        Stmp2 =UNICDbl(lgcTB_47B.GetData(1, "W6"),0) - UNICDbl(lgcTB_47B.GetData(1, "W4"),0)
		     End if
		    IF  UniCDBL(lgcTB_47B.GetData(1, "W7"),0)  <> UNICDBl(lgcTB_47B.GetData(1, "W5"),0) - UNICDbl(lgcTB_47B.GetData(1, "W5"),4)  Or   UniCDBL(lgcTB_47B.GetData(1, "W7"),0) <> Stmp2 Then
		        blnError = True	
		        
			
		       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(1, "W7"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "������"," �������ݾ�_�Ű��� - �׸�(4)ȸ����ݾ� �Ǵ� (�׸�(5)�� �׸�(6)�������ݾ�_���Լ���� �� ū �ݾ�-�׸�(4)ȸ����ݾ�  "))
		    End If
		
		Else
		   blnError = True	
		End If  
		
		
 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W7"), 15, 0)
		

		Call lgcTB_47B.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop

	' 2. �׸���2
	Do Until lgcTB_47B.EOF(2) 
	
		sTmp = lgcTB_47B.GetData(2, "W8")
		
		Select Case sTmp
			Case "108"
				sMsg = "��������"
			Case "109"
				sMsg = "����δ��"
			Case "110"
				sMsg = "��������"
		End Select
		

		
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W9"), sMsg & "_�ݾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W9"), 15, 0)
				
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W10"), sMsg & "_�������ڻ갡��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W11"),  sMsg & "_ȸ��ձݰ���") Then blnError = True	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W11"), 15, 0)
		
		If  ChkNotNull(lgcTB_47B.GetData(2, "W12"), sMsg & "_�ѵ��ʰ���") Then 
		    '- �׸�(11)ȸ��ձݰ��� - �׸�(10)�������ڻ갡�� 
		   IF UniCDBL(lgcTB_47B.GetData(2, "W12"),0)  <>  UNICDBl(lgcTB_47B.GetData(2, "W11"),0)   -UNICDBl(lgcTB_47B.GetData(2, "W10"),0)  Then
		      Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(2, "W12"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_�ѵ��ʰ���", "(11)ȸ��ձݰ��� - �׸�(10)�������ڻ갡��"))
		      blnError = True	
		
		   End If
		Else
		    blnError = True	
		End If    
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W12"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W13"), sMsg & "_�̻����ͱݻ��Ծ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W13"), 15, 0)
		
		Call lgcTB_47B.MoveNext(2)	' -- 2�� ���ڵ�� 
	Loop
	
	' 3. �׸��� 3
	
		
			

	If Not ChkNotNull(lgcTB_47B.GetData(3, "W14"), "�����ޱ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W14"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W15"), "������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W15"), 15, 0)
				
	If  ChkNotNull(lgcTB_47B.GetData(3, "W16"), "����") Then 
	   '���� ����	- �׸�(14)�����ޱ� - �׸�(15)������ 
	    IF UNICDBl(lgcTB_47B.GetData(3, "W16"),0) <> UNICDBl(lgcTB_47B.GetData(3, "W14"),0) -UNICDBl(lgcTB_47B.GetData(3, "W15"),0) Then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(3, "W16"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_����", "�׸�(14)�����ޱ� - �׸�(15)������"))
	
	    End If
	Else
	   blnError = True	
	End If
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W16"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W17"), "��������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W17"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W18"), "ȸ�����") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W18"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W19"), "������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W19"), 15, 0)
				
	
	' 4. �׸���4
	Do Until lgcTB_47B.EOF(4) 
	
		sTmp = lgcTB_47B.GetData(4, "W20")
		
		Select Case sTmp
			Case "111"
				sMsg = "�Ǽ��Ϸ��ڻ��"
			Case "112"
				sMsg = "�Ǽ������ڻ��"
			Case "113"
				sMsg = "��"
		End Select
		
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W21"), sMsg & "_�Ǽ��ڱ�����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W21"), 15, 0)
				
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W22"), sMsg & "_ȸ�����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W22"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W23"),  sMsg & "_�󰢴���ڻ��") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W23"), 15, 0)
		
		''����������_�Ǽ��Ϸ��ڻ��	= �׸�(21)�Ǽ��ڱ����� - �׸�(22)ȸ����� - �׸�(23)�󰢴���ں�		
		If  ChkNotNull(lgcTB_47B.GetData(4, "W24"), sMsg & "_����������") Then 
		    IF UNICDbl(lgcTB_47B.GetData(4, "W24"),0) <> UNIcdbl(lgcTB_47B.GetData(4, "W21"),0)-UNIcdbl(lgcTB_47B.GetData(4, "W22"),0) - UNICDbl(lgcTB_47B.GetData(4, "W23"),0) Then
		       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(3, "W24"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_����", "�׸�(21)�Ǽ��ڱ����� - �׸�(22)ȸ����� - �׸�(23)�󰢴���ں�"))
		       blnError = True	 
		    End IF
		    
		    
		Else
			blnError = True	
		End IF	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W24"), 15, 0)

		Call lgcTB_47B.MoveNext(4)	' -- 2�� ���ڵ�� 
	Loop
	
	sHTFBody = sHTFBody & UNIChar("", 117) & vbCrLf	' -- ���� 

	PrintLog "WriteLine2File : " & sHTFBody

 
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_47B = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9103MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9103MA1 : " & lgStrSQL
End Sub
%>
