<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��52ȣ Ư�������ڰ� �ŷ����� 
'*  3. Program ID           : W9107MA1
'*  4. Program Name         : W9107MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_52

Set lgcTB_52 = Nothing ' -- �ʱ�ȭ 

Class C_TB_52
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
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

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
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_52H	WITH (NOLOCK) " & vbCrLf	' ����52ȣ 
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				If WHERE_SQL = "" Then	' -- �ܺ�ȣ��ܿ� �氪�� �����Ѵ�. �������� 200603�ݿ�: 2/15���ڸ��ϳ������� 
					lgStrSQL = lgStrSQL & "	AND (W3 > 0 OR W10 > 0) " & vbCrLf		' -- �谡 0���� ū�͸� �ݿ� 
					lgStrSQL = lgStrSQL & "	UNION ALL " & vbCrLf
					lgStrSQL = lgStrSQL & "	SELECT 999999 SEQ_NO, '', '', SUM(W3), SUM(W4), SUM(W5), SUM(W6), SUM(W7), SUM(W8), SUM(W9), SUM(W10), SUM(W11), SUM(W12), SUM(W13), SUM(W14), SUM(W15), SUM(W16), MAX(INSRT_USER_ID), MAX(INSRT_DT), MAX(UPDT_USER_ID), MAX(UPDT_DT) " & vbCrLf
					lgStrSQL = lgStrSQL & " FROM TB_52H	WITH (NOLOCK) " & vbCrLf	
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & " HAVING SUM(W3) > 0 OR SUM(W10) > 0 " & vbCrLf	
				End If
				
	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_52H2	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9107MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9107MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9107MA1"

	Set lgcTB_52 = New C_TB_52		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_52.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9107MA1

	'==========================================
	' -- ��52ȣ Ư�������ڰ� �ŷ����� �������� 
	' -- 1. ����׸��԰ŷ��� 
	iSeqNo = 1	
	
	Do Until lgcTB_52.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_52.GetData(1, "SEQ_NO"), 0) = 999999 Then
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(1, "SEQ_NO"), 6)
		Else
			' -- �հ谡 �ƴҶ��� ����: 200603���� 
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		
			If Not ChkNotNull(lgcTB_52.GetData(1, "W1"), "���θ�(����)") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(1, "W1"), 60)
	  
	  
			If  ChkNotNull(lgcTB_52.GetData(1, "W2"), "����ڵ�Ϲ�ȣ(�ֹε�Ϲ�ȣ)") Then 
				If  Len(Replace(lgcTB_52.GetData(1, "W2"),"-","") ) <> 10  and Len(Replace(lgcTB_52.GetData(1, "W2"),"-","") ) <> 13 then 
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W2"), UNIGetMesg(TYPE_CHK_CHARNUM, "����ڵ�Ϲ�ȣ(�ֹε�Ϲ�ȣ)","10 �̰ų� 13"))
					blnError = True	
				End If
			   
			    If lgcCompanyInfo is Nothing Then
					blnError = True	
					Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("���ڽŰ� Conversion���α׷����� �ڵ�A100 ���α��� ���������� üũ�Ǿ�� �մϴ�.", "",""))
					Exit Function
			    End If
			    
				If replace(lgcCompanyInfo.OWN_RGST_NO,"-","") = Replace(lgcTB_52.GetData(1, "W2"),"-","") Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W2"), UNIGetMesg("������ ����ڹ�ȣ�� �ŷ����� ����ڹ�ȣ�� ���� �� �����ϴ�", "",""))
					blnError = True	
				 End If
		
			Else
			   blnError = True	
			End If

		End If
		
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_52.GetData(1, "W2")), 13)
		
		If  ChkNotNull(lgcTB_52.GetData(1, "W3"), "����ŷ���_��") Then
		    '����ŷ���_�� = �׸� (9)����_����ڻ� + (10)����_��Ÿ + (11)�����ڻ� + (12)�뿪+ (13)������� + (14)��Ÿ 
		       sTmp =  UNICDbl(lgcTB_52.GetData(1, "W4"),0) +  UNICDbl(lgcTB_52.GetData(1, "W5"),0) + UNICDbl(lgcTB_52.GetData(1, "W6"),0) + _
		               UNICDbl(lgcTB_52.GetData(1, "W7"),0) +  UNICDbl(lgcTB_52.GetData(1, "W8"),0) +  UNICDbl(lgcTB_52.GetData(1, "W9"),0) 
				If  UNICDbl(lgcTB_52.GetData(1, "W3"),0) <> sTmp Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W3"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_52.GetData(1, "W1") & "�� ����ŷ���_��","����_����ڻ� + ����_��Ÿ + �����ڻ� + �뿪+ ������� + ��Ÿ"))
				     blnError = True	
				End if
		
		    
		Else
			blnError = True	
		End If	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W3"), 15, 0)
				
		If Not ChkNotNull(lgcTB_52.GetData(1, "W4"), "����ŷ���_�����ڻ� ����ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W5"), "����ŷ���_�����ڻ� ��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W5"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W6"), "����ŷ���_�����ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W6"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W7"), "����ŷ���_�뿪") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W7"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W8"), "����ŷ���_�������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W8"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W9"), "����ŷ���_��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W9"), 15, 0)
		
		If  ChkNotNull(lgcTB_52.GetData(1, "W10"), "���԰ŷ���_��") Then
		     '���԰ŷ���_�� = - �׸� (16)����_����ڻ� + (17)����_��Ÿ + (18)�����ڻ� + (19)�뿪+ (20)������� + (21)��Ÿ,0) + _
		       sTmp =  UNICDbl(lgcTB_52.GetData(1, "W11"),0) +  UNICDbl(lgcTB_52.GetData(1, "W12"),0) + UNICDbl(lgcTB_52.GetData(1, "W13"),0) + _
		               UNICDbl(lgcTB_52.GetData(1, "W14"),0) +  UNICDbl(lgcTB_52.GetData(1, "W15"),0) +  UNICDbl(lgcTB_52.GetData(1, "W16"),0) 
				If  UNICDbl(lgcTB_52.GetData(1, "W10"),0) <> sTmp Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_52.GetData(1, "W1") & "�� ���԰ŷ���_��","����_����ڻ� + ����_��Ÿ + �����ڻ� + �뿪+ ������� + ��Ÿ"))
				     blnError = True	
				End if
		Else
			blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W11"), "���԰ŷ���_�����ڻ� ����ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W11"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W12"), "���԰ŷ���_�����ڻ� ��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W12"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W13"), "���԰ŷ���_�����ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W13"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W14"), "���԰ŷ���_�뿪") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W14"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W15"), "���԰ŷ���_�������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W15"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W16"), "���԰ŷ���_��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W16"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 5) & vbCrLf	' -- ���� 

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_52.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop

	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = "" ' -- �ں��ŷ��� �ڿ� ����ؾߵǹǷ� �� �ʱ�ȭ: 200603
	' -- 2. �ں��ŷ� 
	iSeqNo = 1	
	
	Do Until lgcTB_52.EOF(2) 
		sHTFBody = sHTFBody & "83"
		'sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		sHTFBody = sHTFBody & UNIChar("A232", 4)		' -- 200603 �������� 
		
		'If UNICDbl(lgcTB_52.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		'Else
		'	sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "SEQ_NO"), 6)
		'End If
				
		If Not ChkNotNull(lgcTB_52.GetData(2, "W17"), "���θ�(��ȣ �Ǵ� ����)") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W17"), 60)
		
	
		
		If  ChkNotNull(lgcTB_52.GetData(2, "W18"), "����ڵ�Ϲ�ȣ(�ֹε�Ϲ�ȣ)") Then 
			If  Len(Replace(lgcTB_52.GetData(2, "W18"),"-","") ) <> 10  and Len(Replace(lgcTB_52.GetData(2, "W18"),"-","") ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(2, "W18"), UNIGetMesg(TYPE_CHK_CHARNUM, "����ڵ�Ϲ�ȣ(�ֹε�Ϲ�ȣ)","10 �̰ų� 13"))
				blnError = True	
			End If
		   
			If replace(lgcCompanyInfo.OWN_RGST_NO,"-","") = Replace(lgcTB_52.GetData(2, "W18"),"-","") Then
			    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(2, "W18"), UNIGetMesg("������ ����ڹ�ȣ�� �ŷ����� ����ڹ�ȣ�� ���� �� �����ϴ�", "",""))
				blnError = True	
			 End If
		
		Else
		   blnError = True	
		End If
		 
		 
		 sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_52.GetData(2, "W18")), 13)
		
		'If Not ChkNotNull(lgcTB_52.GetData(2, "W19"), "����,���� ����") Then blnError = True
		If lgcTB_52.GetData(2, "W19") = "0" Then
			sHTFBody = sHTFBody & UNIChar("", 1)
		Else
			if Not ChkBoundary("1,2",lgcTB_52.GetData(2, "W19"),"����,���� ����") then   blnError = True	
						
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W19"), 1)
			
			If Not ChkNotNull(lgcTB_52.GetData(2, "W21"), "����,���� ����") Then blnError = True	'������ üũ�� ���ڴ� �ʼ�)
		End If
			
		sHTFBody = sHTFBody & UNI8Date(lgcTB_52.GetData(2, "W21"))
	
		If Not ChkNotNull(lgcTB_52.GetData(2, "W22"), "����(�Ǵ� ����)��_�׸��Ѿ� ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W22"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W23"), "����(�Ǵ� ����)��_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W23"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W24"), "����(�Ǵ� ����)��_�׸��Ѿ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W24"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W25"), "����(�Ǵ� ����)��_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W25"), 5, 2)
		
		'If Not ChkNotNull(lgcTB_52.GetData(2, "W26"), "�պ�,�����պ� ����") Then blnError = True
		If lgcTB_52.GetData(2, "W26") = "0" Then
			sHTFBody = sHTFBody & UNIChar("", 1)
		Else
		   	if Not ChkBoundary("1,2",lgcTB_52.GetData(2, "W26"),"�պ�,�����պ� ����") then   blnError = True	
						
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W26"), 1)
			  
			If Not ChkNotNull(lgcTB_52.GetData(2, "W28"), "�պ�,�����պ� ����") Then blnError = True	'�պ�,�����պ� ����üũ�� ���ڴ� �ʼ�)
		End If	
		
		If lgcTB_52.GetData(2, "W19") = "0" And lgcTB_52.GetData(2, "W26") = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("����/����, �պ�/�����պ� �� ��� 1������ üũ�Ǿ�� �˴ϴ�.", "",""))
		End If

		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_52.GetData(2, "W28"))

		If Not ChkNotNull(lgcTB_52.GetData(2, "W29_1"), "�պ����ε� ���ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W29_1"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W29_2"), "�պ����ε� ����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W29_2"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W30_1"), "���պ����ε� ���ڻ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W30_1"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W30_2"), "���պ����ε� ����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W30_2"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W31"), "�պ�����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W31"), 5, 2)
		
		sHTFBody = sHTFBody & UNIChar("", 12) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_52.MoveNext(2)	' -- 2�� ���ڵ�� 
	Loop
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		'Call Write2File(sHTFBody)	' -- 200603 ���� 
		Call PushRememberDoc(sHTFBody)	' -- �ٷ� ������� �ʰ� ����Ų��(inc_HomeTaxFunc.asp�� ����)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_52 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL
			
			lgStrSQL = ""
			' -- ǥ�ؼ��Ͱ�꼭(A115,A116)�� ��������(�Ϲݹ����� �ڵ�(82) ����.����.���Ǿ������� �ڵ�(73))�� ��ġ���� ������ ���� 
			lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- ǥ�ؼ��Ͱ�꼭 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- ���α���(�Ϲ�/����)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '82'"		 	 & vbCrLf	' -- ���α���(�Ϲ�)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '73'"		 	 & vbCrLf	' -- ���α���(����)
			End If
	End Select
	PrintLog "SubMakeSQLStatements_W9107MA1 : " & lgStrSQL
End Sub
%>
