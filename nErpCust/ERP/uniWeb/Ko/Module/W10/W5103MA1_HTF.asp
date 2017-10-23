<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��15ȣ ���񺰼ҵ�ݾ��������� 
'*  3. Program ID           : W5103MA1
'*  4. Program Name         : W5103MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_15

Set lgcTB_15 = Nothing ' -- �ʱ�ȭ 

Class C_TB_15
	' -- ���̺��� �÷����� 
	
	Dim SELECT_SQL
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

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

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst

	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm)
		End If
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				
				If SELECT_SQL = "" Then
					lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & SELECT_SQL & vbCrLf
				End If
				
	            If WHERE_SQL = "" Then 
					lgStrSQL = lgStrSQL & " , ( SELECT ITEM_NM FROM TB_ADJUST_ITEM WITH (NOLOCK)  WHERE ITEM_CD = A.W1 ) W1_NM " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & " , '' W1_NM  " & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_15	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W5103MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W5103MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W5103MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W5103MA1"

	Set lgcTB_15 = New C_TB_15		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_15.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W5103MA1

	'==========================================
	' -- ��15ȣ �ҵ�ݾ������հ�ǥ �������� 
	iSeqNo = 1	: sHTFBody = ""
	
	Do Until lgcTB_15.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W_TYPE"), 1)
		
		If UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_15.GetData("W3"), lgcTB_15.GetData("W1_NM") & " ó���ڵ�") Then blnError = True	
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("SEQ_NO"), 6)
			iSeqNo = 1	' -- W_TYPE ���� ���� �ʱ�ȭ 
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W1_NM"), 40)
		
		If  Not ChkMinusAmt(lgcTB_15.GetData("W2"),"����üũ") Then blnError = True	' -- ����üũ 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_15.GetData("W2"), 15, 0)
 
		If lgcTB_15.GetData("W_TYPE") = "1" and UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999  Then
			If Not ChkBoundary("100,200,300,400,500,600", lgcTB_15.GetData("W3"), lgcTB_15.GetData("W_TYPE") & " ó���ڵ�") Then blnError = True
		ElseIf lgcTB_15.GetData("W_TYPE") = "2" and UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999 Then
			If Not ChkBoundary("100,200", lgcTB_15.GetData("W3"), lgcTB_15.GetData("W1_NM") & " ó���ڵ�") Then blnError = True
			 
		End If
			
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W3"), 4)
	
		If lgcTB_15.GetData("W_TYPE") = "1" And lgcTB_15.GetData("W3") = "400" Then sT1_400SUM = sT1_400SUM + UNICDbl(lgcTB_15.GetData("W2"), 0)
		If lgcTB_15.GetData("W_TYPE") = "2" And lgcTB_15.GetData("W3") = "100" Then sT2_100SUM = sT2_100SUM + UNICDbl(lgcTB_15.GetData("W2"), 0)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		lgcTB_15.MoveNext 
	Loop
	
	' �ͱݻ��Թ׼ձݺһ����� �ڵ� '400'�� �� �Ǵ� �ձݻ��Թ��ͱݺһ����� �ڵ� '100'�� �� �ݾ��� '0'�� �ƴϸ� �ں��ݰ���������������(��)(A103) ���� �ʼ� �Է� 
	If sT1_400SUM > 0 Or sT2_100SUM > 0 Then
		' -- ��50ȣ �ں��ݰ� ��������������(��)
		Set cDataExists.A103 = new C_TB_50A	' -- W7105MA1_HTF.asp �� ���ǵ� 
			
		' -- �߰� ��ȸ������ �о�´�.
		cDataExists.A103.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A103.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
		If Not cDataExists.A103.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "�ͱݻ��Թ׼ձݺһ����� �ڵ� '400'�� �� �Ǵ� �ձݻ��Թ��ͱݺһ����� �ڵ� '100'�� �� �ݾ��� '0'�� �ƴϸ� �ں��ݰ���������������(��)(A103) ���� �ʼ� �Է� ")		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		End If
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A103 = Nothing
	End If

	' ----------- 
	'Call SubCloseRs(oRs2)

	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_15 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W5103MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W5103MA1 : " & lgStrSQL
End Sub
%>
