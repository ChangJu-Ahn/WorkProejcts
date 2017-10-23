
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��3ȣ ���μ� ����ǥ�� �� ����������꼭 
'*  3. Program ID           : W8101MA1
'*  4. Program Name         : W8101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_3

Set lgcTB_3 = Nothing ' -- �ʱ�ȭ 

Class C_TB_3
	' -- ���̺��� �÷����� 
	Dim W01
	Dim W02
	Dim W03
	Dim W04
	Dim W05
	Dim W54
	Dim W06
	Dim W07
	Dim W08
	Dim W09
	Dim W10
	Dim W11
	Dim W12
	Dim W13
	Dim W14_CD
	Dim W14
	Dim W15
	Dim W16
	Dim W17
	Dim W18
	Dim W19
	Dim W20
	Dim W21
	Dim W22
	Dim W23
	Dim W24
	Dim W25_NM
	Dim W25
	Dim W26
	Dim W27
	Dim W28
	Dim W29
	Dim W30
	Dim W31
	Dim W32
	Dim W33
	Dim W34
	Dim W35_CD
	Dim W35
	Dim W36
	Dim W37
	Dim W38
	Dim W39
	Dim W40
	Dim W41
	Dim W42
	Dim W43_NM
	Dim W43
	Dim W44
	Dim W45
	Dim W46
	Dim W47
	Dim W48
	Dim W49
	Dim W50
	Dim W51
	Dim W52
	Dim W53
	Dim W55
	
	' -- 2005-01-04 : 200603 ���� 
	Dim W55_1
	Dim W56

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs1
				 
		On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		LoadData = False
			 
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W01			= oRs1("W01")
		W02			= oRs1("W02")
		W03			= oRs1("W03")
		W04			= oRs1("W04")
		W05			= oRs1("W05")
		W54			= oRs1("W54")
		W06			= oRs1("W06")
		W07			= oRs1("W07")
		W08			= oRs1("W08")
		W09			= oRs1("W09")
		W10			= oRs1("W10")
		W11			= oRs1("W11")
		W12			= oRs1("W12")
		W13			= oRs1("W13")
		W14_CD			= oRs1("W14_CD")
		W14			= oRs1("W14")
		W15			= oRs1("W15")
		W16			= oRs1("W16")
		W17			= oRs1("W17")
		W18			= oRs1("W18")
		W19			= oRs1("W19")
		W20			= oRs1("W20")
		W21			= oRs1("W21")
		W22			= oRs1("W22")
		W23			= oRs1("W23")
		W24			= oRs1("W24")
		W25_NM			= oRs1("W25_NM")
		W25			= oRs1("W25")
		W26			= oRs1("W26")
		W27			= oRs1("W27")
		W28			= oRs1("W28")
		W29			= oRs1("W29")
		W30			= oRs1("W30")
		W31			= oRs1("W31")
		W32			= oRs1("W32")
		W33			= oRs1("W33")
		W34			= oRs1("W34")
		W35_CD			= oRs1("W35_CD")
		W35			= oRs1("W35")
		W36			= oRs1("W36")
		W37			= oRs1("W37")
		W38			= oRs1("W38")
		W39			= oRs1("W39")
		W40			= oRs1("W40")
		W41			= oRs1("W41")
		W42			= oRs1("W42")
		W43_NM			= oRs1("W43_NM")
		W43			= oRs1("W43")
		W44			= oRs1("W44")
		W45			= oRs1("W45")
		W46			= oRs1("W46")
		W47			= oRs1("W47")
		W48			= oRs1("W48")
		W49			= oRs1("W49")
		W50			= oRs1("W50")
		W51			= oRs1("W51")
		W52			= oRs1("W52")
		W53			= oRs1("W53")
		W55			= oRs1("W55")
		
		' -- 2005-01-04 : 200603 ���� 
		W55_1		= oRs1("W55_1")
		W56			= oRs1("W56")
		
		Call SubCloseRs(oRs1)
		
		LoadData = True
	End Function

	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub	
	
	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_3	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
End Class

' -- �� ���Ŀ��� ������ �ٸ� ���ĵ� 
Class TYPE_DATA_EXIST_W8101MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
	Dim A106
	Dim A159
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8101MA1"

	Set lgcTB_3 = New C_TB_3		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_3.LoadData Then Exit Function			' -- ��3ȣ ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W8101MA1

	' -- �������� 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

	'==========================================
	' -- ��3ȣ ���μ�����ǥ�� �� ����������꼭 �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If ChkNotNull(lgcTB_3.W01, "��꼭���������") Then ' -- ����Ÿ����� ������ 
		If lgcTB_1.W2 <> "50" Then ' -- ���������������� �������Ͱ������� '50' �� �ƴ� ��� 
			
			' -- ��3ȣ3(1)(2)ǥ�ؼ��Ͱ�꼭(A115)���� 
			Set cDataExists.A115 = new C_TB_3_3	' -- W1105MA1_HTF.asp �� ���ǵ� 
			
			' -- �߰� ��ȸ������ �о�´�.
			Call SubMakeSQLStatements_W8101MA1("A115_1",iKey1, iKey2, iKey3)   
			
			cDataExists.A115.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
			cDataExists.A115.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
			If Not cDataExists.A115.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��3ȣ3(1)(2)ǥ�ؼ��Ͱ�꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
				'ǥ�ؼ��Ͱ�꼭(A115,A116)�� ��������(�Ϲݹ����� �ڵ�(82) ����.����.���Ǿ������� �ڵ�(73))����ġ���� ������ ���� 
				If UNICDbl(lgcTB_3.W01, 0) <> UNICDbl(cDataExists.A115.W5, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_3.W01, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��꼭���������","ǥ�ؼ��Ͱ�꼭(A115,A116)�� ��������"))
				End If
			End If
		
			' -- ����� Ŭ���� �޸� ���� 
			Set cDataExists.A115 = Nothing
		End If
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W01, 15, 0)
	
	If ChkNotNull(lgcTB_3.W02, "�ҵ������ݾ�_�ͱݻ���") Then ' -- ����Ÿ����� ������ 
		' -- ��15ȣ ���񺰼ҵ�ݾ���������(A102)���� 
		Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp �� ���ǵ� 
			
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W8101MA1("A102_1",iKey1, iKey2, iKey3)   
			
		cDataExists.A102.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A102.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
		If Not cDataExists.A102.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "��15ȣ ���񺰼ҵ�ݾ���������_�ͱݻ��Թ׼ձݺһ���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
			''�ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �ݾ�(2)�� �հ�� ��ġ 
			If UNICDbl(lgcTB_3.W02, 0) <> UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ҵ������ݾ�_�ͱݻ���","�ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �ݾ�(2)�� �հ�"))
			End If
		End If
		' -- ����� Ŭ���� �޸� ���� 
		'Set cDataExists.A102 = Nothing

		' -- ��3ȣ3(1)(2)ǥ�ؼ��Ͱ�꼭(A115)���� 
		Set cDataExists.A115 = new C_TB_3_3	' -- W1105MA1_HTF.asp �� ���ǵ� 
			
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W8101MA1("A115_2",iKey1, iKey2, iKey3)   
			
		cDataExists.A115.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A115.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
		
		If Not cDataExists.A115.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "��3ȣ3(1)(2)ǥ�ؼ��Ͱ�꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
 
			'ǥ�ؼ��Ͱ�꼭(A115,A116)�� �׸�(81)�� ���μ����(�Ϲݹ���)/�׸�(72)���μ����(��������)���� ������ ���� 
			If UNICDbl(lgcTB_3.W02, 0)< UNICDbl(cDataExists.A115.W5, 0) Then

				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg(TYPE_CHK_LOW_AMT, "�ҵ������ݾ�_�ͱݻ���","ǥ�ؼ��Ͱ�꼭(A115,A116)�� �׸�(81)�� ���μ����(�Ϲݹ���)/�׸�(72)���μ����(��������)"))
			End If
			
		End If	
		
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A115 = Nothing
					
	Else
		blnError = True
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W02, 15, 0)
	
	If  ChkNotNull(lgcTB_3.W03, "�ҵ������ݾ�_�ձݻ���") Then 
		' -- ��15ȣ ���񺰼ҵ�ݾ���������(A102)���� 
		Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp �� ���ǵ� 
			
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W8101MA1("A102_2",iKey1, iKey2, iKey3)   
			
		cDataExists.A102.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A102.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
		If Not cDataExists.A102.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "��15ȣ ���񺰼ҵ�ݾ���������_�ձݻ��Թ��ͱݺһ���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
			'�ҵ�ݾ������հ�ǥ(A102)�� �ձݻ��Թ��ͱݺһ����� �ݾ�(5)�� �հ�� ��ġ 
			If UNICDbl(lgcTB_3.W03, 0) <> UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W03, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ҵ������ݾ�_�ձݻ���","�ҵ�ݾ������հ�ǥ(A102)�� �ձݻ��Թ��ͱݺһ����� �ݾ�(5)�� �հ�"))
			End If
		End If
		' -- ����� Ŭ���� �޸� ���� 
		'Set cDataExists.A102 = Nothing

	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W03, 15, 0)
		
	If Not ChkNotNull(lgcTB_3.W04, "�������ҵ�ݾ�") Then blnError = True	' -- ���α׷��������������Ƿ� �н� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W04, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W05, "��α��ѵ��ʰ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W05, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W54, "��α��ѵ��ʰ� �̿��׼ձݻ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W54, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W06, "��������� �ҵ�ݾ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W06, 15, 0)

	If Not ChkNotNull(lgcTB_3.W07, "�̿���ձ�") Then blnError = True
	If Not ChkMinusAmt(lgcTB_3.W07, "�̿���ձ�") Then blnError = True	' -- ����üũ 
	If UNICDbl(lgcTB_3.W06, 0) < 0 And UNICDbl(lgcTB_3.W07, 0) > 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, "�ڵ�(06)������⵵�ҵ�ݾ��� �����ε� �ݾ��� ������ ����")
	End If
	If UNICDbl(lgcTB_3.W07, 0) > UNICDbl(lgcTB_3.W06, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, UNIGetMesg(TYPE_CHK_LOW_AMT, "��������� �ҵ�ݾ�","�̿���ձ�"))
	End If
	
	If UNICDbl(lgcTB_3.W07, 0) > 0 Then
		' -- 0���� ū ��� A144 �ݵ�� ���� 
		
		' -- �̿���ձ� ��� ���α׷� 
		Set cDataExists.A144 = new C_W7107MA1	' -- W7107MA1_HTF.asp �� ���ǵ�: ���α׷��̶� Ŭ�������� ���α׷�ID
			
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W8101MA1("A144",iKey1, iKey2, iKey3)   
			
		cDataExists.A144.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A144.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
		If Not cDataExists.A144.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "�̿���ձ� ���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
			'�̿���ձ� ��� (6)�������� �հ�� ��ġ 
			If UNICDbl(lgcTB_3.W07, 0) <> UNICDbl(cDataExists.A144.W6, 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�̿���ձ�","�̿���ձ� ���_(6)�������� �հ�"))
			End If
		End If	
		
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A144 = Nothing
	End If

	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W07, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W08, "������ҵ�") Then blnError = True
	If Not ChkMinusAmt(lgcTB_3.W08, "������ҵ�") Then blnError = True	' -- ����üũ 
	
	If UNICDbl(lgcTB_3.W06, 0) < 0 And UNICDbl(lgcTB_3.W08, 0) > 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W08, "�ڵ�(06)������⵵�ҵ�ݾ��� �����ε� �ݾ��� ������ ����")
	End If
	If UNICDbl(lgcTB_3.W08, 0) > (UNICDbl(lgcTB_3.W06, 0) - UNICDbl(lgcTB_3.W07, 0)) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W08, UNIGetMesg(TYPE_CHK_HIGH_AMT, "������ҵ�","�ڵ�(06)������⵵�ҵ�ݾ� - �ڵ�(07)�̿���ձ�"))
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W08, 15, 0)
	
	
	
	If Not ChkNotNull(lgcTB_3.W09, "�ҵ����") Then blnError = True
	
	'2007.03.27 �߰� lws
	'�ҵ���� 0���� ũ�� 
	' <= �׸�(108)����������ҵ�ݾ�-�̿���ձ�(�ڵ�07)-������ҵ�(�ڵ�08) 
	
	if lgcTB_3.W09 > 0 then
	
		if lgcTB_3.W09  <= lgcTB_3.W06 - lgcTB_3.W07 - lgcTB_3.W08 then 'pass
		else
			Call SaveHTFError(lgsPGM_ID, lgcTB_3.W09, UNIGetMesg(TYPE_CHK_HIGH_AMT, "�ҵ����","�׸�(108)����������ҵ�ݾ�-�̿���ձ�(�ڵ�07)-������ҵ�(�ڵ�08) "))
		end if

	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W09, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W10, "����ǥ�رݾ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W10, 15, 0)
	
	If Not ChkNotNull(UNICDBl(lgcTB_3.W11,0) * 100, "����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(UNICDBl(lgcTB_3.W11,0) * 100, 5, 2)

	If Not ChkNotNull(lgcTB_3.W12, "���⼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W12, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W13, "���������ҵ�") Then blnError = True
	If UNICDbl(lgcTB_3.W13, 0) <> 0 And lgcTB_1.W1 = "2" Then
		If Not SearchTaxDocCd("A162") Then	' wa101mb2.asp �� ���ǵ� 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_3.W13, "�ܱ������̰� ���������ҵ��� '0'�� �ƴѰ�� ���������ҵ�ݾװ�꼭(A162) ����")
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W13, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W14, "����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W14, 5, 2)
	
	If Not ChkNotNull(lgcTB_3.W15, "���⼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W15, 15, 0)
 
	If Not ChkNotNull(lgcTB_3.W16, "���⼼���հ�") Then blnError = True
	If UNICDbl(lgcTB_3.W16, 0) <> UNICDbl(lgcTB_3.W15, 0)+ UNICDbl(lgcTB_3.W12, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W16, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���⼼���հ�","���⼼��"))
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W16, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W17, "�������鼼��(��)") Then blnError = True
	' 0 ���� ū ��� A106 üũ (8ȣ��)
	If UNICDbl(lgcTB_3.W17, 0) > 0 Then
		' -- ��8ȣ(��)��������...
		
		' --- ����Ÿ ��ȸ SQL
		Call SubMakeSQLStatements_W8101MA1("A106",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs2,lgStrSQL, "", "") = False Then
			blnError = True
		    Call SaveHTFError(lgsPGM_ID, "��8ȣ(��)�������鼼�׹��߰����μ����հ�ǥ", TYPE_DATA_NOT_FOUND)
		Else
		
			If UNICDbl(lgcTB_3.W17, 0) <> UNICDbl(oRs2("W20"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W17, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������鼼��(��)","��8ȣ(��)�������鼼�׹��߰����μ����հ�ǥ"))
			End If
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W17, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W18, "��������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W18, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W19, "�������鼼��(��)") Then blnError = True
	If UNICDbl(lgcTB_3.W19, 0) > UNICDbl(lgcTB_3.W18, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W19, UNIGetMesg(TYPE_CHK_HIGH_AMT, "�������鼼��(��)","��������"))
	End If	
	If UNICDbl(lgcTB_3.W19, 0) > 0 Then
		If Not oRs2 is Nothing Then	' -- ������ �ε�Ǿ��ٸ�..
			If UNICDbl(lgcTB_3.W19, 0) <> UNICDbl(oRs2("W21"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W19, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������鼼��(��)","��8ȣ(��)�������鼼�׹��߰����μ����հ�ǥ"))
			End If
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W19, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W20, "���꼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W20, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W21, "������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W21, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W22, "���ѳ����μ���_�߰���������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W22, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W23, "���ѳ����μ���_���úΰ�����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W23, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W24, "���ѳ����μ���_��õ���μ���") Then blnError = True

	' 2006-01-04 : 200603 ���� 
	'If UNICDbl(lgcTB_3.W24, 0) <> 0 Then
		' -- ��10ȣ 
		Set cDataExists.A159 = new C_TB_10A	' -- W7101MA1_HTF.asp �� ���ǵ� 
			
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W8101MA1("A159",iKey1, iKey2, iKey3)   
			
		cDataExists.A159.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A159.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
			
		If Not cDataExists.A159.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "��10ȣ ��õ���μ��׸���", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
			'��10ȣ ��õ���μ��׸��� (6)���μ��հ�� ��ġ 
			If UNICDbl(lgcTB_3.W24, 0) <> UNICDbl(cDataExists.A159.W6, 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W24, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���ѳ����μ���_��õ���μ���","��10ȣ ��õ���μ��׸���_(6)���μ��հ�"))
			End If
		End If	
		
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A159 = Nothing
	'End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W24, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W25, "���ѳ����μ���_��������ȸ����� �ܱ����μ���") Then blnError = True
	
	If UNICDbl(lgcTB_3.W25, 0) <> 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W25, UNIGetMesg("�ڵ�(25)�� ���ѳ����μ���_��������ȸ����� �ܱ����μ����� 0���� ū ��� �������ڵ��� �ܱ����μ��� ��꼭(���� ��10ȣ�� 2 ����)�� �־���մϴ�. �ش缭���� �������� ������ uniERP ���� �����Ͻñ� �ٶ��ϴ�", "",""))
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W25, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W26, "���ѳ����μ���_�Ұ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W26, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W27, "�Ű�����������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W27, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W28, "�ⳳ�μ����հ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W28, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W29, "������߰����μ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W29, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W30, "�����Ҽ��װ��_������������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W30, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W31, "������絵�ҵ濡���ѹ��μ����_�絵����_����ڻ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W31, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W32, "������絵�ҵ濡���ѹ��μ����_�絵����_�̵���ڻ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W32, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W33, "������ҵ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W33, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W34, "����ǥ��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W34, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W35, "����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W35, 5, 2)
	
	If Not ChkNotNull(lgcTB_3.W36, "���⼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W36, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W37, "���鼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W37, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W38, "��������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W38, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W39, "��������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W39, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W40, "���꼼��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W40, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W41, "������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W41, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W42, "�ⳳ�μ���_���úΰ�����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W42, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W43, "�ⳳ�μ���_(" & lgcTB_3.W43_NM & ")����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W43, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W44, "�ⳳ�μ���_��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W44, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W45, "���������Ҽ���_������絵�ҵ濡���ѹ��μ����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W45, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W46, "���װ�_���������Ҽ��װ�") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W46, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W55, "���װ�_��ǰ��ٸ�ȸ��ó���������װ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W55, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W47, "���װ�_�г����װ�������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W47, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W48, "�г��Ҽ���_���ݳ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W48, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W49, "�г��Ҽ���_����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W49, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W50, "�г��Ҽ���_��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W50, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W51, "�������μ���_���ݳ���") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W51, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W52, "�������μ���_����") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W52, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W53, "�������μ���_��") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W53, 15, 0)

	' -- 2006-01-04 : 200603 ���� 
	If Not ChkNotNull(lgcTB_3.W55_1, "����ǥ������") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W55_1, 15, 0)

	If UNICDbl(lgcTB_3.W55_1, 0) > 0 Then	' -- ����ǥ�������� 0���� ū ��� ����ǥ������ �������(A224)�� �׸�(7)����ǥ������ �ݾװ� ��ġ ��(�̰���)
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W55_1, "����ǥ�������� 0���� ū ��� ����ǥ������ �������(A224)�� �׸�(7)����ǥ������ �ݾװ� ��ġ ��")
	End If
	
	If UNICDbl(lgcTB_3.W56, 0) <> (UNICDbl(lgcTB_3.W10, 0) + UNICDbl(lgcTB_3.W55_1, 0)) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W56, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(56)����ǥ�رݾ�","�ڵ�(10)����ǥ�رݾ� + �ڵ�(55)����ǥ������"))
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W56, 15, 0)

	'sHTFBody = sHTFBody & UNIChar("", 69)	' -- ���� 
	sHTFBody = sHTFBody & UNIChar("", 19)	' -- ���� : 2006-01-05 : 200603 ������ 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	
 
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3 = Nothing	' -- �޸����� 

End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8101MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A115_1" '-- �ܺ� ���� SQL
			
			lgStrSQL = ""
			' -- ǥ�ؼ��Ͱ�꼭(A115,A116)�� ��������(�Ϲݹ����� �ڵ�(82) ����.����.���Ǿ������� �ڵ�(73))�� ��ġ���� ������ ���� 
			'lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- ǥ�ؼ��Ͱ�꼭 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- ���α���(�Ϲ�/����)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '82'"		 	 & vbCrLf	' -- ���α���(�Ϲ�)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '73'"		 	 & vbCrLf	' -- ���α���(����)
			End If

	  Case "A115_2" '-- �ܺ� ���� SQL
			
			lgStrSQL = ""
			' -- ǥ�ؼ��Ͱ�꼭(A115,A116)�� �׸�(81)���μ����(�Ϲݹ���)/�׸�(72)���μ����(��������) 
			'lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- ǥ�ؼ��Ͱ�꼭 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- ���α���(�Ϲ�/����)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '81'"		 	 & vbCrLf	' -- ���α���(�Ϲ�)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '72'"		 	 & vbCrLf	' -- ���α���(����)
			End If
			
	  Case "A102_1" '-- �ܺ� ���� SQL
		
			' -- �ҵ�ݾ������հ�ǥ(A102)�� �ͱݻ��Թ׼ձݺһ����� �ݾ�(2)�� �հ�� ��ġ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '1'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf

	  Case "A102_2" '-- �ܺ� ���� SQL
	  
			lgStrSQL = ""
			' -- �ҵ�ݾ������հ�ǥ(A102)�� �ձݻ��Թ��ͱݺһ����� �ݾ�(5)�� �հ�� ��ġ 
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '2'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf
			
	  Case "A144" '-- �ܺ� ���� SQL
	  
			lgStrSQL = ""
			' -- �̿���ձݵ�� 
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf	
	  
	  Case "A106"
			lgStrSQL = ""
			' -- ��8ȣ(��) �������鼼��	
			lgStrSQL = lgStrSQL & "SELECT " & vbCrLf
			lgStrSQL = lgStrSQL & " ISNULL( ( " & vbCrLf
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf	
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '50' " & vbCrLf	
			lgStrSQL = lgStrSQL & "	), 0) - ISNULL( (" & vbCrLf	
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf	' 
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '66' " & vbCrLf
			lgStrSQL = lgStrSQL & "	), 0) AS W20 " & vbCrLf	
			lgStrSQL = lgStrSQL & ",ISNULL( ( " & vbCrLf
			lgStrSQL = lgStrSQL & "		SELECT W4 " & vbCrLf					' -- �׸���1 (10)�Ұ� 
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '0' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '10' " & vbCrLf	
			lgStrSQL = lgStrSQL & "	), 0) + ISNULL( (" & vbCrLf	
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf					' -- (66) �����η�..(�����Ѽ�..)		
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '66' " & vbCrLf
			lgStrSQL = lgStrSQL & "	), 0) AS W21 " & vbCrLf		
			lgStrSQL = lgStrSQL & " " & vbCrLf	

	  Case "A159" '-- �ܺ� ���� SQL
	  
			lgStrSQL = ""
			' -- �̿���ձݵ�� 
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf	
	  
	End Select
	PrintLog "SubMakeSQLStatements_W8101MA1 : " & lgStrSQL
End Sub

%>
