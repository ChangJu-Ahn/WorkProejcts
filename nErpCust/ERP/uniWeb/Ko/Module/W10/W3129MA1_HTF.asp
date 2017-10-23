<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��20ȣ �����󰢺���� �հ�ǥ 
'*  3. Program ID           : W3129MA1
'*  4. Program Name         : W3129MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_20

Set lgcTB_20 = Nothing ' -- �ʱ�ȭ 

Class C_TB_20
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	
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

			
		' ��Ƽ�������� ù���� ���� 
		Call GetData
		
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
		Call GetData()
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
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
				lgStrSQL = lgStrSQL & " FROM TB_20	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W3129MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W3129MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum
    Dim chkW1_1,chkW1_2,chkW1_3, chkW1_4, chkW1_5, chkW1_6, chkW1_7,chkData,chkW1
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W3129MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W3129MA1"

	Set lgcTB_20 = New C_TB_20		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_20.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W3129MA1

	'==========================================
	' -- ��20ȣ �����󰢺���� �հ�ǥ �������� 
	iSeqNo = 1	: sHTFBody = ""
	 
	 chkW1_1 = "�⸻�����"
	 chkW1_2 = "�����󰢴����"
	 chkW1_3 = "�̻��ܾ�"
	 chkW1_4 = "�󰢹�����"
	 chkW1_5 = "ȸ��ձݰ���"
	 chkW1_6 = "�󰢺��ξ�"
	 chkW1_7 = "���κ�����"

	
	Do Until lgcTB_20.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If Not ChkNotNull(lgcTB_20.W1, "�ڻ걸���ڵ�") Then blnError = True	
		chkW1 = ""
		sTmp = lgcTB_20.W1
		Select Case sTmp
			Case "1"
				 chkW1 = chkW1_1         '�޼������� �����ڵ�� ���� �ʰ� ���� �ֱ� ���� 
			     chkW1_1 = ""
			Case "2"
				 chkW1 = chkW1_2
			     chkW1_2 = ""
			Case "3"
				 chkW1 = chkW1_3
			     chkW1_3 = ""
			Case "4"
				 chkW1 = chkW1_4
				 chkW1_4 = ""
			Case "5"
				 chkW1 = chkW1_5
				 chkW1_5 = ""
			Case "6"
				 chkW1 = chkW1_6
				 chkW1_6 = ""
			Case "7"
				 chkW1 = chkW1_7	
			     chkW1_7 = ""
		End Select
       
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_20.W1, 1)
		
		
		
		'���հ�� : ����⹰ + ������ġ + ���Ÿ�ڻ� + �칫�������ڻ� 
		if  UNICDbl(lgcTB_20.W2, 0) <>  UNICDbl(lgcTB_20.W3,0) +  UNICDbl(lgcTB_20.W4, 0) +  UNICDbl(lgcTB_20.W5, 0) + UNICDbl(lgcTB_20.W6,  0) then
		   	Call SaveHTFError(lgsPGM_ID, lgcTB_20.W2, UNIGetMesg(TYPE_CHK_NOT_EQUAL,chkW1 & "�� �հ��","(3)���๰ + (4)�����ġ + (5)��Ÿ�ڻ� + (6)���������ڻ� "))
		   	 blnError = True	
		end if
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W2, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W3, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W4, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W5, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W6, 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 18) & vbCrLf	' -- ���� 
	  
		lgcTB_20.MoveNext 
	Loop
	
	
	'�ڻ걸���ڵ尡 ���ڵ� ���� ��� ���;� �Ѵ�.
	'( ���ڵ尡 7���� ���� �Ǿ�� �Ѵ�.)

	if chkW1_1 <> "" Or chkW1_2 <> "" or chkW1_3<>"" Or chkW1_4<>"" Or chkW1_5<>"" Or chkW1_6<>""  Or chkW1_7<>"" then
	    chkData =chkW1_1 & " " & chkW1_2  &  " " & chkW1_3 &  " " & chkW1_4 & " " & chkW1_5 &  " " & chkW1_6 &  " " & chkW1_7
	   	Call SaveHTFError(lgsPGM_ID,Trim(chkData), UNIGetMesg(TYPE_CHK_NULL, chkData))
	   	 blnError = True	
	end if  
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_20 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W3129MA1(pMode, pCode1, pCode2, pCode3)
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
	PrintLog "SubMakeSQLStatements_W3129MA1 : " & lgStrSQL
End Sub
%>
