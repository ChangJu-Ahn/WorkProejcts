<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��50ȣ �ں��ݰ� ��������������(��)
'*  3. Program ID           : W7109MA1
'*  4. Program Name         : W7109MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_50B

Set lgcTB_50B = Nothing ' -- �ʱ�ȭ 

Class C_TB_50B
	' -- ���̺��� �÷����� 

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
		Call GetData()
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
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
	            If WHERE_SQL = "" Then 
					lgStrSQL = lgStrSQL & " , ( SELECT ITEM_NM FROM TB_ADJUST_ITEM WITH (NOLOCK)  WHERE ITEM_CD = A.W1 ) W1_NM " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & " , '' W1_NM  " & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_50B	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W7109MA1
	Dim A102

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W7109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo,sTmp1,sTmp2
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7109MA1"

	Set lgcTB_50B = New C_TB_50B		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_50B.LoadData Then Exit Function			' -- ��50ȣ �ں��ݰ� ��������������(��) ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W7109MA1

	'==========================================
	' -- ��15ȣ �ҵ�ݾ������հ�ǥ �������� 
	iSeqNo = 1	: sHTFBody = ""
	
	Do Until lgcTB_50B.EOF() 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_50B.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_50B.GetData("W1_NM"), "���� �Ǵ� ����") Then blnError = True	
			
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("SEQ_NO"), 6)
			
			' -- �հ��϶� �����ǽ� 
			' -- ���� 15ȣ�� �ε��Ѵ�.
			Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp �� ���ǵ� 
				
			' -- �߰� ��ȸ������ �о�´�.  400�� 
			Call SubMakeSQLStatements_W7109MA1("A102_1",iKey1, iKey2, iKey3)   
				
			cDataExists.A102.CALLED_OUT	= True				' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A102.WHERE_SQL	= lgStrSQL			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			cDataExists.A102.SELECT_SQL	= " W3, SUM(W2) W2 "' -- �ٸ� ���� ���� 
			
			
			If Not cDataExists.A102.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��15ȣ ���񺰼ҵ�ݾ���������", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				
		   '�ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��� ó���ڵ� ��400���� �ݾ��� - �ձݻ��� ó���ڵ� ��100���� �ݾ��� = �׸�(4)����� ������ �� - �׸�(3)����� ������ ��		
				'- �ҵ�ݾ������հ�ǥ(A102)�� �ձݻ��� ó���ڵ� ��100:�������� �ݾ��� 
				sTmp1 =  UNICDbl(cDataExists.A102.GetData("W2"), 0) 
				
				Call cDataExists.A102.MoveNext 
				
				'- �ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��� ó���ڵ� ��400:�������� �ݾ��� 
				sTmp2 =  UNICDbl(cDataExists.A102.GetData("W2"), 0) 
				If UNICDbl(lgcTB_50B.GetData("W4"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0)  <> UNICDbl(sTmp2, 0) -  UNICDbl(sTmp1, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_50B.GetData("W4"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0)  , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����� ������ �� - ����� ������ ��","�ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��� ó���ڵ� ��400���� �ݾ��� - �ձݻ��� ó���ڵ� ��100���� �ݾ���"))
				End If
			End If
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A102 = Nothing

		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("W1_NM"), 40)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W2"), lgcTB_50B.GetData("W1_NM") & "_�����ܾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W2"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W3"), lgcTB_50B.GetData("W1_NM") & "_����߰���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W3"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W4"), lgcTB_50B.GetData("W1_NM") & "_���������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W5"), lgcTB_50B.GetData("W1_NM") & "_�⸻�ܾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W5"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("W_DESC"), 50) ' -- ����� 
	
		' -- �⸻�ܾ� = �����ܾ� - ����߰��� + ��������� 
		If UNICDbl(lgcTB_50B.GetData("W2"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0) + UNICDbl(lgcTB_50B.GetData("W4"), 0) <> UNICDbl(lgcTB_50B.GetData("W5"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�⸻�ܾ�","�����ܾ� - ����߰��� + ���������"))
		End If
	
		sHTFBody = sHTFBody & UNIChar("", 38) & vbCrLf	' -- ���� 
		
		iSeqNo = iSeqNo + 1
		
		lgcTB_50B.MoveNext 
	Loop

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_50B = Nothing	' -- �޸����� 

End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W7109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A102_1" '-- �ܺ� ���� SQL
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND (  (A.W_TYPE	= '1' AND A.W3 = '400')  " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		OR (A.W_TYPE	= '2' AND A.W3 = '100') )" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	GROUP BY A.W3 " 	 & vbCrLf

	End Select
	PrintLog "SubMakeSQLStatements_W7109MA1 : " & lgStrSQL
End Sub
%>
