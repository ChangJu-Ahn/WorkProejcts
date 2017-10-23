<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ������1ȣ���󰡰ݻ���� 
'*  3. Program ID           : W9117MA1
'*  4. Program Name         : W9117MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_KJ1

Set lgcTB_KJ1 = Nothing	' -- �ʱ�ȭ 

Class C_TB_KJ1
	' -- ���̺��� �÷����� 
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
			 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		LoadData = False
			 
		PrintLog "LoadData IS RUNNING: "
			 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		lgStrSQL = ""
		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements
		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly)  = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If	
			LoadData = False  
			Exit Function  
		End If

		
		LoadData = True
	End Function

	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function

	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	End Function	
		
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "  A.W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10"
            lgStrSQL = lgStrSQL & " FROM TB_KJ1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	    AND A.W1 <> '4'"  & vbCrLf
			

			
			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

            lgStrSQL = lgStrSQL & " ORDER BY  A.W1 ASC" & vbcrlf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9117MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9117MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows, iSeqNo
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9117MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9117MA1"
	
	Set lgcTB_KJ1 = New C_TB_KJ1		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_KJ1.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ������1ȣ���󰡰ݻ���� �� �������� 
	
	iSeqNo = 1

	For iDx = 2 To 10
		
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' �Ϸù�ȣ 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx),stmp & "����Ư��������_���θ�(��ȣ)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 60)
		stmp =lgcTB_KJ1.GetData("W" & iDx)
	
		lgcTB_KJ1.MoveNext
		lgcTB_KJ1.MoveNext  
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx),stmp & "����Ư��������_���籹��") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 3)
		
		lgcTB_KJ1.MoveNext 
			
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "����Ư��������_��ǥ��(����)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 30)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "����Ư��������_����") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 7)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_KJ1.GetData("W" & iDx), stmp & "����Ư��������_�Ű��ΰ��ǰ���") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 1)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "����Ư��������_������") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 70)
		
		sHTFBody = sHTFBody & UNIChar("", 17) & vbCrLf	' -- ���� 
	
		lgcTB_KJ1.MoveFirst

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' �������� �ƴҶ�, ���� ������ ����Ż�� 
			If lgcTB_KJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	
	Next
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	
	blnError = False : sHTFBody = "" : lgcTB_KJ1.MoveFirst 
	
	' -- �ֽļ�������Ȳ 
	iSeqNo = 1	

	For iDx = 2 To 10
        Call lgcTB_KJ1.Find("W1='1'")	' ���θ� : ������ �����ϱ� ���� : 2006.03.06 �ּ����߰� 
        stmp =lgcTB_KJ1.GetData("W" & iDx)
        
		Call lgcTB_KJ1.Find("W1='9'")	' ���ŷ����� : 2006.03.06 : ȭ�鿡 �ڵ�/�ڵ������ �и��Ǹ鼭 ���� 

		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' �Ϸù�ȣ 
		
		If Not ChkBoundary("01,02,03,04,05,06,07,08,09,10", UNISeqNo(lgcTB_KJ1.GetData("W" & iDx),2), stmp  & "���ŷ�") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 2)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkBoundary("01,02,03,04,05,06", lgcTB_KJ1.GetData("W" & iDx), stmp  & "���󰡰ݻ�����") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 2)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp  & "���ǹ����������") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 34) & vbCrLf	' -- ���� 
	
		lgcTB_KJ1.MoveFirst 

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' �������� �ƴҶ�, ���� ������ ����Ż�� 
			If lgcTB_KJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	
	Next
	
		PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_KJ1 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9117MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_W9117MA1 : " & lgStrSQL
End Sub

%>
