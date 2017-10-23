<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��64ȣ �����󰢹���Ű� 
'*  3. Program ID           : W9115MA1
'*  4. Program Name         : W9115MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_63

Set lgcTB_63 = Nothing ' -- �ʱ�ȭ 

Class C_TB_63
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
		Call SubMakeSQLStatements("A",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("B",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

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
			Case 2
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
				lgStrSQL = lgStrSQL & " FROM TB_63H	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_63A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
			
			Case "B"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
	            
	            If WHERE_SQL = "" Then	' �ܺ�ȣ���� �ƴϸ� ���������� ���������(â����) �ҷ��´� 
					lgStrSQL = lgStrSQL & " , B.FOUNDATION_DT " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " FROM TB_63B	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				
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
Class TYPE_DATA_EXIST_W9115MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9115MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9115MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9115MA1"

	Set lgcTB_63 = New C_TB_63		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_63.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9115MA1

	'==========================================
	' -- ��64ȣ �����󰢹���Ű� �������� 
	' -- 1. ����׸��԰ŷ��� 
	iSeqNo = 1	
	
	If lgcTB_63.EOF(2) Then
			' -- �׸��尡 ���������ʴ��ٸ� ��� 1�� ���� 
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
			
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
					
			sHTFBody = sHTFBody & UNIChar("", 30)
			
			sHTFBody = sHTFBody & UNIChar("", 2)
			sHTFBody = sHTFBody & UNIChar("", 2)
			
			sHTFBody = sHTFBody & UNINumeric(0, 2, 0)
			sHTFBody = sHTFBody & UNINumeric(0, 2, 0)
						
			sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- ���� 
	Else
	
		Do Until lgcTB_63.EOF(2) 
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
			
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			'	���뿬������, �Ű��뿬��, ���泻�뿬�� ��� ��0���̸� ���� ����	
			If Not ChkNotNull(lgcTB_63.GetData(2, "W8"), "�ڻ� �� ������") Then 
			
			    if UNINumeric(lgcTB_63.GetData(2, "W9_Fr"), 2, 0)  <> 0 Or UNINumeric(lgcTB_63.GetData(2, "W9_To"), 2, 0) <> 0 Or  UNINumeric(lgcTB_63.GetData(2, "W10"), 2, 0)  <> 0 Or UNINumeric(lgcTB_63.GetData(2, "W11"), 2, 0) <> 0 then
			       blnError = True	
			    End if   
			End if    
			sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(2, "W8"), 30)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W9_Fr"), "���뿬������_From") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W9_Fr"), 2, 0)
					
			If Not ChkNotNull(lgcTB_63.GetData(2, "W9_To"), "���뿬������_To") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W9_To"), 2, 0)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W10"), "�Ű��뿬��") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W10"), 2, 0)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W11"), "���泻�뿬��") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W11"), 2, 0)
			
			
			sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(2, "W12"), 50)
			
			sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- ���� 

			iSeqNo = iSeqNo + 1
			
			Call lgcTB_63.MoveNext(2)	' -- 1�� ���ڵ�� 
		Loop
	End If
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	' -- 2. �ں��ŷ� 
	
	Do Until lgcTB_63.EOF(3) 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

		If  ChkNotNull(lgcTB_63.GetData(3, "FOUNDATION_DT"), "���������") Then 
		    if ChkDate(lgcTB_63.GetData(3, "FOUNDATION_DT") , "���������") = False  Then blnError = True	
		Else
		    blnError = True	
		End if    
		sHTFBody = sHTFBody & UNI8Date(lgcTB_63.GetData(3, "FOUNDATION_DT"))
		
		If Not ChkNotNull(lgcTB_63.GetData(3, "W7"), "����������������") Then blnError = True	
		sHTFBody = sHTFBody & UNI8Date(lgcTB_63.GetData(3, "W7"))
		
		If Trim(lgcTB_63.GetData(3, "W13_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W13_A")) <> null  Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W13_A"), "���������ڻ�_�Ű�󰢹��") Then blnError = True
		Else
		 blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_A"), 1)
	
	
		
		If  Trim(lgcTB_63.GetData(3, "W13_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W13_B")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W13_B"), "���������ڻ�_����󰢹��") Then blnError = True
		Else
		 blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_B"), 1)
		
		
		'���������ڻ�_������� 
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_C"), 50)
		
		
	
		If  Trim(lgcTB_63.GetData(3, "W14_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W14_A")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W14_A"), "������_�Ű�󰢹����") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_A"), 1)
		
	
		If  Trim(lgcTB_63.GetData(3, "W14_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W14_B")) <> null    Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W14_B"), "������_����󰢹��") Then blnError = True
		Else
		   blnError = True	
		End if 
		   sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_B"), 1)
		

		'������_������� 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_C"), 50)
		
	
		If Trim(lgcTB_63.GetData(3, "W15_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W15_A")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W15_A"), "����������ڻ�_�Ű�󰢹��") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W15_A"), 1)
		

		If  Trim(lgcTB_63.GetData(3, "W15_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W15_B")) <> null Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W15_B"), "����������ڻ�_����󰢹��") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(1, "W15_B"), 1)

        '����������ڻ�_������� 
	
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(1, "W15_C"), 50)
	
		if Trim(lgcTB_63.GetData(3, "W13_A")) = "" and  Trim(lgcTB_63.GetData(3, "W14_A")) = "" and Trim(lgcTB_63.GetData(3, "W15_A")) = ""  then
		   Call SaveHTFError(lgsPGM_ID, Trim(lgcTB_63.GetData(3, "W13_A")), UNIGetMesg(TYPE_CHK_NULL, " �׸�(13) ���������ڻ�_�Ű�󰢹��, �׸�(14) ������_�Ű�󰢹��,�׸�(15) ����������ڻ�_�Ű�󰢹�� �� �ϳ�",""))
		   blnError = True	
		end if
 	
		sHTFBody = sHTFBody & UNIChar("", 22) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_63.MoveNext(3)	' -- 2�� ���ڵ�� 
	Loop
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_63 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9115MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9115MA1 : " & lgStrSQL
End Sub
%>
