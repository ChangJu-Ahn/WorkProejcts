<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : �ؿ��������� ���� 
'*  3. Program ID           : W9127MA1
'*  4. Program Name         : W9127MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : leewolsan
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_A126

Set lgcTB_A126 = Nothing ' -- �ʱ�ȭ 

Class C_TB_A126
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
	            lgStrSQL = lgStrSQL & " A.*  , B.W7, B.W8" & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A126	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_A125	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO=B.SEQ_NO " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W01 > 0 AND ISNULL(B.W7,'') <>''"  & vbCrLf	' -- ����Ÿ�� ���� ���� 
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9127MA1
	Dim A126
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9127MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), sHTFBody2
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    
    PrintLog "MakeHTF_W9127MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9127MA1"

	Set lgcTB_A126 = New C_TB_A126		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_A126.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9127MA1

	' -- �������� 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

	'==========================================
	' -- ��4ȣ �����Ѽ�������꼭 �������� 
	iSeqNo = 1	: sHTFBody = ""

	Do Until lgcTB_A126.EOF 

		' -------------- �繫��Ȳǥ 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & "A126"		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
        '3 �Ϸù�ȣ 
		If Not ChkNotNull(lgcTB_A126.GetData("SEQ_NO"), "�Ϸù�ȣ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("SEQ_NO"), 6, 0)
			
	'4 �������θ� 

		If Not ChkNotNull(lgcTB_A126.GetData("W7"), "�������θ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A126.GetData("W7"), 60)

		'5 �������ΰ�����ȣ 
		If Not ChkNotNull(lgcTB_A126.GetData("W8"), "�������ΰ�����ȣ") Then blnError = True
		' -- 2006.03.29 ����  = 8 ���� 
		' -- ù���ڰ� 1,2 �� �ƴϸ� 
		If lgcTB_A126.GetData("W8") <> "99999999" And (Left(lgcTB_A126.GetData("W8"), 1) <> "1" And Left(lgcTB_A126.GetData("W8"), 1) <> "2"  And Left(lgcTB_A126.GetData("W8"), 1) <> "8") Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A126.GetData("W8"), UNIGetMesg("�������ΰ�����ȣ�� 99999999�� �ƴҶ�, ù���ڰ� 1 �Ǵ� 2 �Ǵ� 8 ��(��) �ƴϸ� �����Դϴ�", "",""))
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_A126.GetData("W8"), 8)
		
		'6.�ڻ��Ѱ� 
		
		If Not ChkNotNull(lgcTB_A126.GetData("W01"), "�ڻ��Ѱ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W01")), "�ڻ��Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W01"), 15, 0)
		
		
		'7.����ä��(Ư��������)
	
		If Not ChkNotNull(lgcTB_A126.GetData("W03"), "����ä��(Ư��������)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W03")), "����ä��(Ư��������)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W03"), 15, 0)

		'8 ����ä��(��Ÿ)
		If Not ChkNotNull(lgcTB_A126.GetData("W04"), "����ä��(��Ÿ)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W04")), "����ä��(��Ÿ)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W04"), 15, 0)
		
		'9 ����ڻ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W05"), "����ڻ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W05")), "����ڻ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W05"), 15, 0)
		
		'10 �������� 
		If Not ChkNotNull(lgcTB_A126.GetData("W06"), "��������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W06")), "��������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W06"), 15, 0)
		
		
		'11 �뿩��(Ư��������)
		If Not ChkNotNull(lgcTB_A126.GetData("W07"), "�뿩��(Ư��������)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W07")), "�뿩��(Ư��������)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W07"), 15, 0)
		
		'12 �뿩��(��Ÿ)
		If Not ChkNotNull(lgcTB_A126.GetData("W08"), "�뿩��(��Ÿ)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W08")), "�뿩��(��Ÿ)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W08"), 15, 0)
		
		'13 �����ڻ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W09"), "�����ڻ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W09")), "�����ڻ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W09"), 15, 0)
		
		'14�����װ��๰ 
		If Not ChkNotNull(lgcTB_A126.GetData("W10"), "�����װ��๰") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W10")), "�����װ��๰") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W10"), 15, 0)
		
		'15 �����ġ,������ݱ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W11"), "�����ġ,������ݱ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W11")), "�����ġ,������ݱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W11"), 15, 0)



		'16 �����ڻ� ��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W12"), "�����ڻ� ��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W12")), "�����ڻ� ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W12"), 15, 0)


		'17 �����ڻ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W13"), "�����ڻ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W13")), "�����ڻ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W13"), 15, 0)
		
		'18 �ڻ��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W14"), "�ڻ� ��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W14")), "�ڻ� ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W14"), 15, 0)
		
		'19 ��ä�Ѱ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W15"), "��ä�Ѱ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W15")), "��ä�Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W15"), 15, 0)
		
		'20 ����ä��(Ư��������)
		If Not ChkNotNull(lgcTB_A126.GetData("W16"), "����ä��(Ư��������)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W16")), "����ä��(Ư��������)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W16"), 15, 0)
		
		'21 ����ä��(��Ÿ)
		If Not ChkNotNull(lgcTB_A126.GetData("W17"), "����ä��(��Ÿ)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W17")), "����ä��(��Ÿ)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W17"), 15, 0)


		'22 ���Ա�(Ư��������)
		If Not ChkNotNull(lgcTB_A126.GetData("W18"), "���Ա�(Ư��������)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W18")), "���Ա�(Ư��������)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W18"), 15, 0)


		'23 ���Ա�(��Ÿ)
		If Not ChkNotNull(lgcTB_A126.GetData("W19"), "���Ա�(��Ÿ)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W19")), "�ڻ��Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W19"), 15, 0)
		
		'24 �����ޱ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W20"), "�����ޱ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W20")), "�����ޱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W20"), 15, 0)
		
		'25 ��ä��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W21"), "��ä ��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W21")), "��ä ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W21"), 15, 0)
		
		'26 �ں����Ѱ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W22"), "�ں����Ѱ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W22")), "�ں����Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W22"), 15, 0)
		
		'27 �ں��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W23"), "�ں���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W23")), "�ں���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W23"), 15, 0)
		
		'28 ��Ÿ�ں��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W24"), "��Ÿ�ں���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W24")), "��Ÿ�ں���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W24"), 15, 0)
		
		'29 �ں��׿��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W25"), "�ں��׿���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W25")), "�ں��׿���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W25"), 15, 0)
		
		'30 �����׿��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W26"), "�����׿���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W26")), "�����׿���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W26"), 15, 0)
		
		'31 ��Ÿ�ں��� �� ��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W27"), "��Ÿ�ں��� �� ��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W27")), "��Ÿ�ں��� �� ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W27"), 15, 0)
		
		'32 ����� 
		If Not ChkNotNull(lgcTB_A126.GetData("W28"), "�����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W28")), "�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W28"), 15, 0)
		
		'33 ���� 
		If Not ChkNotNull(lgcTB_A126.GetData("W29"), "����� ����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W29")), "����� ����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W29"), 15, 0)
		
		'34 ����ױ�Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W30"), "����� ��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W30")), "����� ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W30"), 15, 0)
		
		'35 ������� 
		If Not ChkNotNull(lgcTB_A126.GetData("W31"), "�������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W31")), "�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W31"), 15, 0)
		
		
		'36 �Ǹź�Ͱ����� 
		If Not ChkNotNull(lgcTB_A126.GetData("W34"), "�Ǹź�� �Ϲݰ�����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W34")), "�Ǹź�� �Ϲݰ�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W34"), 15, 0)
		
		'37 �޿�(��ȸ���İ�����)
		If Not ChkNotNull(lgcTB_A126.GetData("W35"), "�޿�(�����İ�����)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W35")), "�޿�(�����İ�����)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W35"), 15, 0)
		
		'38 �޿�(��Ÿ)
		If Not ChkNotNull(lgcTB_A126.GetData("W36"), "�޿�(��Ÿ)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W36")), "�޿�(��Ÿ)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W36"), 15, 0)
		
		
		'39 ������ 
		If Not ChkNotNull(lgcTB_A126.GetData("W37"), "������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W37")), "������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W37"), 15, 0)
		
		'40 �������ߺ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W38"), "�������ߺ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W38")), "�������ߺ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W38"), 15, 0)
		
		'41 ��ջ󰢺� 
		If Not ChkNotNull(lgcTB_A126.GetData("W39"), "��ջ󰢺�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W39")), "��ջ󰢺�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W39"), 15, 0)
		
		'42 �Ǹź�� �Ϲݰ�����-��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W40"), "�Ǹź�� �Ϲݰ�����_��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W40")), "�Ǹź�� �Ϲݰ�����_��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W40"), 15, 0)
		
		'43 �����ܼ��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W41"), "�����ܼ���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W41")), "�����ܼ���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W41"), 15, 0)
		
		'44 ���ڼ��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W42"), "���ڼ���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W42")), "���ڼ���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W42"), 15, 0)
		
		'45 ������ 
		If Not ChkNotNull(lgcTB_A126.GetData("W43"), "������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W43")), "������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W43"), 15, 0)
		
		'46 �����ܼ���-��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W44"), "�����ܼ���_��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W44")), "�����ܼ���_��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W44"), 15, 0)
		
		
		'47 �����ܺ�� 
		If Not ChkNotNull(lgcTB_A126.GetData("W45"), "�����ܺ��") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W45")), "�����ܺ��") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W45"), 15, 0)
		
		
		'48 ���ں�� 
		If Not ChkNotNull(lgcTB_A126.GetData("W46"), "���ں��") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W46")), "���ں��") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W46"), 15, 0)
		
		
		'49 �����ܺ��-��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W47"), "�����ܺ��_��Ÿ") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W47")), "�����ܺ��_��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W47"), 15, 0)
		
		'50 Ư������ 
		If Not ChkNotNull(lgcTB_A126.GetData("W48"), "Ư������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W48")), "Ư������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W48"), 15, 0)
		
		'51 Ư���ս� 
		If Not ChkNotNull(lgcTB_A126.GetData("W51"), "Ư���ս�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W51")), "Ư���ս�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W51"), 15, 0)
		
		'52 ���μ� 
		If Not ChkNotNull(lgcTB_A126.GetData("W52"), "���μ�") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W52")), "���μ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W52"), 15, 0)
		
		'53 �������� 
		If Not ChkNotNull(lgcTB_A126.GetData("W53"), "��������") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W53")), "��������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W53"), 15, 0)
		
		'54 ���ݰ����� 
		If Not ChkNotNull(lgcTB_A126.GetData("W02"), "���ݰ�����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W02")), "���ݰ�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W02"), 15, 0)
		
		
		'55 �������-�������κ��͸��� 
		If Not ChkNotNull(lgcTB_A126.GetData("W32"), "�������-�������κ��͸���") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W32")), "�������-�������κ��͸���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W32"), 15, 0)
		
		
		'56 �������- ��Ÿ���� 
		If Not ChkNotNull(lgcTB_A126.GetData("W33"), "�������- ��Ÿ����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W33")), "�������- ��Ÿ����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W33"), 15, 0)
		
		'57 Ư������ - ä�������� 
		If Not ChkNotNull(lgcTB_A126.GetData("W49"), "�������- ��Ÿ����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W49")), "�������- ��Ÿ����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W49"), 15, 0)
		
		'57 Ư������ - ��Ÿ 
		If Not ChkNotNull(lgcTB_A126.GetData("W50"), "�������- ��Ÿ����") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W50")), "�������- ��Ÿ����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W50"), 15, 0)

		'58 ���� 25
		sHTFBody = sHTFBody & UNIChar("", 25) 
		
		If Not blnError Then
			Call WriteLine2File(sHTFBody)
		End If
		sHTFBody=""
		lgcTB_A126.MoveNext 
	Loop


	PrintLog "WriteLine2File : " & sHTFBody
	
	' -- ���Ͽ� ����Ѵ�.

	If Not blnError Then

		'Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_A126 = Nothing	' -- �޸����� 

End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9127MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A126" '-- �ܺ� ���� SQL

			lgStrSQL = ""
			

	End Select
	PrintLog "SubMakeSQLStatements_W9127MA1 : " & lgStrSQL
End Sub
%>
