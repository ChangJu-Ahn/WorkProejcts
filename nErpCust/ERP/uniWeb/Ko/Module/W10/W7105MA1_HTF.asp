<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��50ȣ �ں��ݰ� ��������������(��)
'*  3. Program ID           : W7105MA1
'*  4. Program Name         : W7105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_50A
dim lgcTB_3_2_2
dim lgcCOMPANY_HISTORY

Set lgcTB_50A = Nothing ' -- �ʱ�ȭ 
Set lgcTB_3_2_2 = Nothing ' -- �ʱ�ȭ 

'---------------------------------------------------------------------
'---------------------------------------------------------------------

Class C_TB_50A
	' -- ���̺��� �÷����� 
	Dim W_CD
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W_DESC
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
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
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W_CD		= lgoRs1("W_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W_DESC		= lgoRs1("W_DESC")
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
				lgStrSQL = lgStrSQL & " FROM TB_50A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND NOT( W_CD BETWEEN '171' AND '177' )" '200703 TEMP
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class



'---------------------------------------------------------------------
'---------------------------------------------------------------------





  
Class C_TB_3_2_2
	' -- ���̺��� �÷����� 

	Dim W1 
	Dim W2
	Dim W3
	Dim W4

	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3,iKey4
				 
		On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		if lgcCOMPANY_HISTORY.COMP_TYPE2="1" then '�Ϲ� 
			iKey4="41,44,50,57"
		else
			iKey4="68,69,73,79"
		end if

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3,iKey4)                                       '�� : Make sql statements

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
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
		
			W1			= lgoRs1("W1") '�ں��� 
			W2			= lgoRs1("W2") '�ں��׿��� 
			W3			= lgoRs1("W3") '�����׿��� 
			W4			= lgoRs1("W4") '�ں����� 
			
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
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3,pCode4)
		dim tmpKey
		tmpKey = split(pCode4,",")
	    Select Case pMode 
	      Case "H"

				
				lgStrSQL = ""
	           
				lgStrSQL = lgStrSQL & " SELECT                       "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD =  " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(0)&"'                "
				lgStrSQL = lgStrSQL & "   AND W2 LIKE '3%' ) W1      "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR =" & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(1)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W2       "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(2)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W3     "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " (SELECT A.W6-A.W5 FROM            "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(3)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W4      "

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class







' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W7105MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
End Class



' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W7105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, sMsg
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7105MA1"
	
	Set lgcTB_50A = New C_TB_50A		' -- �ش缭�� Ŭ���� 
	Set lgcTB_3_2_2= New C_TB_3_2_2		' -- �ش缭�� Ŭ���� 
	Set lgcCOMPANY_HISTORY= new C_COMPANY_HISTORY
	
	call lgcCOMPANY_HISTORY.LoadData
	
	If Not lgcTB_50A.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	If Not lgcTB_3_2_2.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 

	'==========================================
	' -- ��3ȣ ���μ�����ǥ�� �� ����������꼭 �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	Do Until lgcTB_50A.EOF 
	
		Select Case lgcTB_50A.W_CD
			Case "01"
				sMsg = "1.�ں���"
			Case "02"
				sMsg = "2.�ں��׿���"
			Case "03"
				sMsg = "3.�����׿���"
			Case "04"
				sMsg = "4.�ں�����"
			Case "05"
				sMsg = ""
			Case "06"
				sMsg =""
			Case "07"
				sMsg = ""
			Case "08"
				sMsg = ""
			Case "09"
				sMsg = ""
			Case "10"
				sMsg = ""
			Case "11"
				sMsg = ""
			Case "12"
				sMsg =""
			Case "13"
				sMsg = ""
			Case "20"
				sMsg = "5.��(I)"
			Case "21"
				sMsg = "6.�ںбݰ������ݰ�꼭(��)��(II)"
			Case "22"
				sMsg = "7.���μ�"
			Case "23"
				sMsg = "8.�ֹμ�"
			Case "30"
				sMsg = "9.��(III)"
			Case "31"
				sMsg = "10.��������(I+II-III)"			
		End Select
	
		Select Case lgcTB_50A.W_CD
			Case "16","17"
				sHTFBody = sHTFBody & UNIChar(lgcTB_50A.W1, 30)	' Null ��� 
		End Select
		
		If Not ChkNotNull(lgcTB_50A.W2, sMsg & "_�����ܾ�") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W2, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W3, sMsg & "_���������_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W3, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W4, sMsg & "_���������_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W4, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W5, sMsg & "_�⸻�ܾ�") Then blnError = True	
		
		'200703 add
		
		
		
		Select Case lgcTB_50A.W_CD
			Case "01" '�ں��� 
		
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W1) then
					
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W1) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"ǥ�ش�������ǥ-�ں���", sMsg & "_�⸻�ܾ�"))
				
				end if
				
			Case "02" '�ں��׿��� 
			
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W2) then
					Call SaveHTFError(lgsPGM_ID,cDbl(lgcTB_3_2_2.W2) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"ǥ�ش�������ǥ-�ں��׿���", sMsg & "_�⸻�ܾ�"))
				end if
				
			Case "14" '�����׿��� 
			
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W3) then
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W3) & "<>" &cDbl(lgcTB_50A.W5)  , UNIGetMesg(TYPE_CHK_NOT_EQUAL,"ǥ�ش�������ǥ-�����׿���", sMsg & "_�⸻�ܾ�"))
				end if
			
			Case "15" '�ں����� 
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W4) then
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W4) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"ǥ�ش�������ǥ-�ں�����", sMsg & "_�⸻�ܾ�"))
				end if
				
				
			case else 
			
		end Select 

		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W5, 15, 0)

		lgcTB_50A.MoveNext 
	Loop
	
	
	sHTFBody = sHTFBody & UNIChar("", 64)	' -- ���� 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_50A = Nothing	' -- �޸����� 
	
End Function


%>
