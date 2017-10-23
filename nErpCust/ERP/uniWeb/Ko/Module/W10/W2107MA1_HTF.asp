
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��17ȣ �����ļ��Աݾ׸��� 
'*  3. Program ID           : W2107MA1
'*  4. Program Name         : W2107MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_17

Set lgcTB_17 = Nothing ' -- �ʱ�ȭ 

Class C_TB_17
	' -- ���̺��� �÷����� 
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
		blnData1 = True : blnData2 = True	: blnData3 = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' --������ �о�´�.
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

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

	'----------- ��Ƽ �� ���� ------------------------
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
	
	Function MoveFirst(Byval pType)
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
				lgStrSQL = lgStrSQL & " FROM TB_17H A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_17_D1 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "     AND  ((A.W4 <> 0  and  A.W3 <>'') OR A.CODE_NO IN ( '11','99')) "  & vbCrLf

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_17_D2 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "     AND A.W9 <> 0  "  & vbCrLf

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W2107MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
	Dim A134
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W2107MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, iSeqNo
    Dim dblAmtW4, dblAmtW5, dblAmtW6, dblAmtW7
    Dim dblAmtSumW4, dblAmtSumW5, dblAmtSumW6, dblAmtSumW7
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W2107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W2107MA1"

	Set lgcTB_17 = New C_TB_17		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_17.LoadData Then Exit Function			
	
	Set cDataExists = new TYPE_DATA_EXIST_W2107MA1
	
	'==========================================
	' --��17ȣ �����ļ��Աݾ׸��� �������� 
	iSeqNo = 1	
	
	Do Until lgcTB_17.EOF(1)
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

		If UNICDbl(lgcTB_17.GetData(1,"CODE_NO"), 0) <> "99" Then
			 
			 sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 
			 dblAmtW4 =  dblAmtW4 + UNICDbl(lgcTB_17.GetData(1,"W4"),0)
		     dblAmtW5 =  dblAmtW5 + UNICDbl(lgcTB_17.GetData(1,"W5"),0)
		     dblAmtW6 =  dblAmtW6 + UNICDbl(lgcTB_17.GetData(1,"W6"),0)
		     dblAmtW7 =  dblAmtW7 + UNICDbl(lgcTB_17.GetData(1,"W7"),0)
		     
		     Response.Write "dblAmtW4=" & dblAmtW4 & vbCrLf
		Else
			  sHTFBody = sHTFBody & UNIChar("999999", 6)
			 
			  dblAmtSumW4 =   UNICDbl(lgcTB_17.GetData(1,"W4"),0)
		      dblAmtSumW5 =   UNICDbl(lgcTB_17.GetData(1,"W5"),0)
		      dblAmtSumW6 =   UNICDbl(lgcTB_17.GetData(1,"W6"),0)
		      dblAmtSumW7 =   UNICDbl(lgcTB_17.GetData(1,"W7"),0)
		      '�պ� 
		      
		      Response.Write "dblAmtSumW4=" & dblAmtSumW4 & vbCrLf
		End If
		
		
	if lgcTB_17.GetData(1,"W3") <> "" And UNICDbl(lgcTB_17.GetData(1,"W4"),0) <> 0 Then	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W1"), "����")					Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W2"), "����")					Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W3"), "����(�ܼ�)�������ȣ") Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W4"), "���Աݾ�_��")			Then blnError = True
			
		If Not ChkNotNull(lgcTB_17.GetData(1,"W5"), "���Աݾ�_����")		Then blnError = True	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W6"), "���Աݾ�_���Ի�ǰ")	Then blnError = True	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W7"), "���Աݾ�_����")		Then blnError = True	
	End If	
			
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W1"), 30)			'���� 
				
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W2"), 30)			'���� 
				
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W3"), 7)			'����(�ܼ�)�������ȣ 
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W4"), 15, 0)	'���Աݾ�_�� = �׸�(5)��������ǰ + �׸� (6)���Ի�ǰ + �׸�(7)���� 

	
		If  UNICDbl(lgcTB_17.GetData(1,"W4"),0) <> UNICDbl(lgcTB_17.GetData(1,"W5"),0) + UNICDbl(lgcTB_17.GetData(1,"W6"),0) +UNICDbl(lgcTB_17.GetData(1,"W7"),0)  Then
			 blnError = True
			 Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(1,"W4"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�_��","�׸�(5)��������ǰ + �׸� (6)���Ի�ǰ + �׸�(7)���� "))		
		End If
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W5"), 15, 0)	'���Աݾ�_���� 
			
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W6"), 15, 0)	'���Աݾ�_���Ի�ǰ 
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W7"), 15, 0)	'���Աݾ�_���� 
		 
		
		sHTFBody = sHTFBody & UNIChar("", 61) & vbCrLf	' -- ���� 
		
		
		lgcTB_17.MoveNext(1) 
		iSeqNo = iSeqNo + 1
	Loop

	If  dblAmtW4 <> dblAmtSumW4  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW4 & " <> " & dblAmtSumW4 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�_�հ�(�ڵ� 99)","�� ���Աݾ�_���� ��"))		
	End If
	
	If  dblAmtW5 <> dblAmtSumW5  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW5 & " <> " & dblAmtSumW5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�_����(�ڵ� 99)","�� ���Աݾ�_������ ��"))		
	End If
	
	If  dblAmtW6 <> dblAmtSumW6  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW6 & " <> " & dblAmtSumW6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�_���Ի�ǰ(�ڵ� 99)","�� ���Աݾ�_���Ի�ǰ�� ��"))		
	End If
	
	If  dblAmtW7 <> dblAmtSumW7  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW7 & " <> " & dblAmtSumW7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�_����(�ڵ� 99)","���Աݾ�_������ ��"))		
	End If

	' -- 200603 ����:  �����ļ��Աݾ׸��� (A111)������ �Ϸù�ȣ�� 999999�϶�   - ���Աݾ���������(A134)�� �׸�(6)�����ļ��Աݾ�_��� ��ġ���� �߰� 
	Set cDataExists.A134 = new C_TB_16	' -- W6127MA1_HTF.asp �� ���ǵ� 
											
	' -- �߰� ��ȸ������ �о�´�.
	cDataExists.A134.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
	cDataExists.A134.WHERE_SQL = " AND A.SEQ_NO = 999999 "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
											
	If Not cDataExists.A134.LoadData() Then
	
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_��󼼾��� '0'���� ū ��� �����Ѽ������꼭(A140) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
	Else

		If UNICDbl(dblAmtSumW4, 0)  <> UNICDbl(cDataExists.A134.GetData(1, "W6") , 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, dblAmtSumW4 & " <> " & cDataExists.A134.GetData(1, "W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��17ȣ �����ļ��Աݾ׸���(A111)������ �ڵ�(99)�հ�","��16ȣ ���Աݾ���������(A134)�� �׸�(6)�����ļ��Աݾ�_��"))
		End If
													
	End If

	' -- ����� Ŭ���� �޸� ���� 
	Set cDataExists.A134 = Nothing		

	'------ ������ ���Աݾ׸���_�ΰ���ġ ����ǥ�ذ� ���Աݾ����װ��� (���)
	iSeqNo = 1
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_17.GetData(2,"W8"), "�ΰ���ġ��_����ǥ��_��") Then blnError = True	
    
    ' -- 200603 : ������ 
	'sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W8"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W9"), "�ΰ���ġ��_����ǥ��_�Ϲ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W9"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W10"), "�ΰ���ġ��_����ǥ��_������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W10"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W11"), "�鼼������Աݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W11"), 15, 0)
		
	If  ChkNotNull(lgcTB_17.GetData(2,"W12"), "�հ�") Then 
	    '�׸�(12)�հ� : �׸�(8)�ΰ�������ǥ��_�� + �׸�(11)�鼼������Աݾ� 
	    IF UNICDbl(lgcTB_17.GetData(2,"W12"),0) <> UNICDbl(lgcTB_17.GetData(2,"W9"),0)  + UNICDbl(lgcTB_17.GetData(2,"W10"),0) + UNICDbl(lgcTB_17.GetData(2,"W11"),0)  Then
	       Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(2,"W12"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�հ�","�׸�(8)�ΰ�������ǥ��_�� + �׸�(11)�鼼������Աݾ�"))		
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End If    
	   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W12"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W13"), "���Աݾ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W13"), 15, 0)

	' -- 2006.03.21 ���� : 1. ������ ���Աݾ׸���(A111)���� �߰� 
	If  UNICDbl(lgcTB_17.GetData(2,"W13"), 0) <> dblAmtSumW4  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,lgcTB_17.GetData(2,"W13") & " <> " & dblAmtSumW4 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "1. ������ ���Աݾ׸����� ���Աݾ�_�հ�(�ڵ� 99)","2. �ΰ���ġ�� ����ǥ�ذ� ���Աݾ� ���װ����� (12) ���Աݾ�"))		
	End If

		
	If  ChkNotNull(lgcTB_17.GetData(2,"W14"), "����") Then 
	    '�׸�(14)���� : �׸�(12)�հ� - �׸�(13)���Աݾ� 
	    IF UNICDbl(lgcTB_17.GetData(2,"W14"),0) <> UNICDbl(lgcTB_17.GetData(2,"W12"),0)  - UNICDbl(lgcTB_17.GetData(2,"W13"),0)  Then
	       Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(2,"W12"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�׸�(14)����","�׸�(12)�հ� - �׸�(13)���Աݾ�"))		
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W14"), 15, 0)

	' -- 200603 ����		
	sHTFBody = sHTFBody & UNIChar("", 54) & vbCrLf	' -- ���� 


	' --��17ȣ �����ļ��Աݾ׸��� - ���Աݾװ��� ���׳��� : 200603 ���� 
	iSeqNo = 1	
	
	Do Until lgcTB_17.EOF(3)
		If UNICDbl(lgcTB_17.GetData(3,"W9"), 0) > 0 Then
	
			sHTFBody = sHTFBody & "85"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
			If Trim(lgcTB_17.GetData(3,"W8")) = "���װ�" Then
				sHTFBody = sHTFBody & "999999"
			Else
				sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			End If
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W8"), "���Աݾ�����_����") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(3,"W8"), 20)			
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W15"), "���Աݾ�����_�ڵ�") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(3,"W15"), 2, 0)	
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W15"), "���Աݾ�����_�ݾ�") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(3,"W9"), 15, 0)	
		
			'If Not ChkNotNull(lgcTB_17.GetData(3,"W_REMARK"), "���Աݾ�����_���") Then blnError = True	' 2006.03.06 ���� 
			sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(3,"W_REMARK"), 20)			
		
			sHTFBody = sHTFBody & UNIChar("", 31) & vbCrLf	' -- ���� 
		
			iSeqNo = iSeqNo + 1
		End If
		
		lgcTB_17.MoveNext(3) 
		
	Loop
	

	' ----------- 
	PrintLog "WriteLine2File : 33" & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_17 = Nothing	' -- �޸����� 
	
End Function


%>
