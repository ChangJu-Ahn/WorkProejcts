<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��������Ư�������ڿ����Ͱ�꼭 
'*  3. Program ID           : W9121MA1
'*  4. Program Name         : W9121MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_KJ_BJ1

Set lgcTB_KJ_BJ1 = Nothing	' -- �ʱ�ȭ 

Class C_TB_KJ_BJ1
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
			LoadData = False  
			Exit Function
		End If

		
		LoadData = True
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
			lgStrSQL = " SELECT  * FROM (" & vbCrLf
			
            lgStrSQL = lgStrSQL & " SELECT CONVERT(INT, A.W1 ) W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10" & vbCrLf
            
			lgStrSQL = lgStrSQL & "  From  dbo.ufn_TB_KJ_BJ1_HOME_TAX_GetRef("& pCode1 &","& pCode2 &","& pCode3 &") A" & vbCrLf
			lgStrSQL = lgStrSQL & "  Union All " & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT  " & vbCrLf
            lgStrSQL = lgStrSQL & "  CONVERT(INT, A.W1 ) W1 ,  Cast(A.W2 as Varchar(15)), Cast(A.W3 as Varchar(15)), Cast(A.W4 as Varchar(15)), Cast(A.W5 as Varchar(15)), " & vbCrLf
            lgStrSQL = lgStrSQL & "   Cast(A.W6 as Varchar(15)),Cast(A.W7 as Varchar(15)), Cast(A.W8 as Varchar(15)), Cast(A.W9 as Varchar(15))  ,  Cast(A.W10 as Varchar(15)) " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_KJ_BJ1 A WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & " ) X "	 & vbCrLf

			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

            lgStrSQL = lgStrSQL & " ORDER BY  X.W1 ASC" & vbcrlf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9121MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9121MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows, iSeqNo
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    Dim dblW01, dblW02, dblW03
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9121MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9121MA1"
	
	Set lgcTB_KJ_BJ1 = New C_TB_KJ_BJ1		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_KJ_BJ1.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ��������Ư�������ڿ����Ͱ�꼭 �� �������� 
	
	iSeqNo = 1

	For iDx = 2 To 10
		
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' �Ϸù�ȣ 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "��Ī") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 40)
	
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "������_�ּ�") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 70)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�������(����)") Then blnError = True		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_KJ_BJ1.GetData("W" & iDx))
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�������(����)") Then blnError = True		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_KJ_BJ1.GetData("W" & iDx))
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�־��� �ڵ�") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 7)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�ں��ݾ� �Ǵ� ���ڱݾ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "Ư�������� ����") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 1)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�ֽĵ��� ��������_���� ��") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�ֽĵ��� ��������_��������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�ֽĵ��� ��������_�Ǽ��� ��") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�ֽĵ��� ��������_�Ǽ�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		lgcTB_KJ_BJ1.MoveNext ' -- ���������� �ǳʶ�: 2006.02.28 �ֿ��¼���(����)
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW01 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW02 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�����Ѽ���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW03 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		If dblW03 <> (dblW01 - dblW02) Then
			Call SaveHTFError(lgsPGM_ID, dblW03, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"�����Ѽ���", "����� - �����"))
		End If
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "�Ǹź�Ͱ�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "��������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "���μ��������������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 25) & vbCrLf	' -- ���� 
	
		lgcTB_KJ_BJ1.MoveFirst 

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' �������� �ƴҶ�, ���� ������ ����Ż�� 
			If lgcTB_KJ_BJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	Next
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_KJ_BJ1 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9121MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_W9121MA1 : " & lgStrSQL
End Sub

%>
