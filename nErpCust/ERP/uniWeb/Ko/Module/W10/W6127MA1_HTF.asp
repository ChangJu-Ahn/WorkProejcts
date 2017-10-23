<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��4ȣ �����Ѽ�������꼭 
'*  3. Program ID           : W6127MA1
'*  4. Program Name         : W6127MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_4

Set lgcTB_4 = Nothing ' -- �ʱ�ȭ 

Class C_TB_4
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
				lgStrSQL = lgStrSQL & " FROM TB_4	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6127MA1
	Dim A103
	Dim A101	' -- ���μ�����ǥ�ؼ���������꼭(A101)
	Dim A108	' -- Ư�������������(A108) - W4105MA1
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6127MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25)
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6127MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6127MA1"

	Set lgcTB_4 = New C_TB_4		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_4.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W6127MA1

	' -- �������� 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

	'==========================================
	' -- ��4ȣ �����Ѽ�������꼭 �������� 
	iSeqNo = 1	: sHTFBody = ""

	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	Do Until lgcTB_4.EOF 
		' -- 2006-01-05 : 200603 ������ 
		arrVal(2, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W2"), 0)
		arrVal(3, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W3"), 0)
		arrVal(4, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W4"), 0)
		arrVal(5, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W5"), 0)
	
'1		�ڷᱸ�� 
'2		�����ڵ� 




		Select Case lgcTB_4.GetData("W1")
			Case "01"
			'3		��꼭���������_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "��꼭���������_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "02"
			'4		�ͱݻ���_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�ͱݻ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "03"
			'5		�ձݻ���_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�ձݻ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "04"
			'6		�����ļҵ�ݾ�_�����ļ��� 
			'7		�����ļҵ�ݾ�_�����Ѽ� 
			'8		�����ļҵ�ݾ�_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�����ļҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�����ļҵ�ݾ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�����ļҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "05"
			'9		�غ��_�����Ѽ� 
			'10		�غ��_������ 
			'11		�غ��_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�غ��_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "�غ��_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�غ��_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "06"
			'12		Ư����_�����Ѽ� 
			'13		Ư����_������ 
			'14		Ư����_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "Ư����_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "Ư����_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "Ư����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "07"
			'15		Ư�����ձݻ����� �ҵ�ݾ�_�����ļ��� 
			'16		Ư�����ձݻ����� �ҵ�ݾ�_�����Ѽ� 
			'17		Ư�����ձݻ����� �ҵ�ݾ�_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "Ư�����ձݻ������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "Ư�����ձݻ������ҵ�ݾ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "Ư�����ձݻ������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "08"
			'18		��α��ѵ��ʰ���_�����ļ��� 
			'19		��α��ѵ��ʰ���_�����Ѽ� 
			'20		��α��ѵ��ʰ���_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "��α��ѵ��ʰ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "��α��ѵ��ʰ���_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "��α��ѵ��ʰ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "09"
			
			'21		��α��ѵ��ʰ��̿��� �ձݻ���_�����ļ��� 
			'22		��α��ѵ��ʰ��̿��� �ձݻ���_�����Ѽ� 
			'23		��α��ѵ��ʰ��̿��� �ձݻ���_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "��α��ѵ��ʰ��̿���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "��α��ѵ��ʰ��̿���_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "��α��ѵ��ʰ��̿���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "10"
			'24		������⵵�ҵ�ݾ�_�����ļ��� 
			'25		������⵵�ҵ�ݾ�_�����Ѽ� 
			'26		������⵵�ҵ�ݾ�_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "����������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "����������ҵ�ݾ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "����������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "11"
			'27		�̿���ձ�_�����ļ��� 
			'28		�̿���ձ�_�����Ѽ� 
			'29		�̿���ձ�_�����ļ��� 


				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�̿���ձ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�̿���ձ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�̿���ձ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "12"
			'30		������ҵ�_�����ļ��� 
			'31		������ҵ�_�����Ѽ� 
			'32		������ҵ�_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "������ҵ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "������ҵ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "������ҵ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			
			Case "13"
			'33		�����Ѽ������������ҵ�_�����Ѽ�  
			'34		�����Ѽ������������ҵ�_������    
			'35		�����Ѽ������������ҵ�_�����ļ��� 

				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�����Ѽ������������ҵ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "�����Ѽ������������ҵ�_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�����Ѽ������������ҵ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "14"
			'36		�����Ѽ��������ͱݺһ���_�����Ѽ�  
			'37		�����Ѽ��������ͱݺһ���_������    
			'38		�����Ѽ��������ͱݺһ���_�����ļ���			

				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�����Ѽ��������ͱݺһ���_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�����Ѽ��������ͱݺһ���_������ ") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�����Ѽ��������ͱݺһ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "15"	
			'39		�������ҵ�ݾ�_�����ļ���            
			'40		�������ҵ�ݾ�_�����Ѽ�              
			'41		�������ҵ�ݾ�_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�������ҵ�ݾ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�������ҵ�ݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "16"

			'42		�ҵ����_�����ļ���                  
			'43		�ҵ����_�����Ѽ�                    
			'44		�ҵ����_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "�ҵ����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�ҵ����_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�ҵ����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
				
			Case "17"
			'45		�����Ѽ������� �ҵ����_�����Ѽ�   
			'46		�����Ѽ������� �ҵ����_������     
			'47		�����Ѽ������� �ҵ����_�����ļ��� 		
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "�����Ѽ�������ҵ����_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "�����Ѽ�������ҵ����_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "�����Ѽ�������ҵ����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "18"
			'48		����ǥ�رݾ�_�����ļ���              
			'49		����ǥ�رݾ�_�����Ѽ�                
			'50		����ǥ�رݾ�_�����ļ���      
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "����ǥ�رݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "����ǥ�رݾ�_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "����ǥ�رݾ�_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "19"
			'51		����_�����ļ���                      
			'52		����_�����Ѽ�                        
			'53		����_�����ļ���         
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 5, 2)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "����_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 5, 2)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "����_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 5, 2)
				
			Case "20"
			'54		���⼼��_�����ļ���                  
			'55		���⼼��_�����Ѽ�                    
			'56		���⼼��_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "���⼼��_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "���⼼��_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"),  15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "���⼼��_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "21"
			'57		���鼼��_�����ļ���                  
			'58		���鼼��_������                      
			'59		���鼼��_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "���鼼��_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "���鼼��_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "���鼼��_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
		
			Case "22"
			'60		���װ���_�����ļ���                  
			'61		���װ���_������                      
			'62		���װ���_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "���װ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "���װ���_������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "���װ���_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)	
			Case "23"
			'63		��������_�����ļ���                  
			'64		��������_�����ļ��� 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "��������_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "��������_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "24"
			'65		����ǥ������_�����ļ���              
			'66		����ǥ������_�����Ѽ�                
			'67		����ǥ������_�����ļ���  
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "����ǥ������_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "����ǥ������_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "����ǥ������_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "25"
			'68		����ǥ�رݾ�2_�����ļ���             
			'69		����ǥ�رݾ�2_�����Ѽ�               
			'70		����ǥ�رݾ�2_�����ļ���   
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "����ǥ�رݾ�2_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "����ǥ�رݾ�2_�����Ѽ�") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "����ǥ�رݾ�2_�����ļ���") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
		End Select
	
		lgcTB_4.MoveNext 
	Loop

	sHTFBody = sHTFBody & UNIChar("", 54) ' 200703 ������ 

	' -- 2006-01-05 : 200603 ������ 		
	' -- ��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)���� 
	Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp �� ���ǵ� 
					
	' -- �߰� ��ȸ������ �о�´�.
	Call SubMakeSQLStatements_W6127MA1("A101",iKey1, iKey2, iKey3)   
					

	if 	arrVal(2, 23) = arrVal(3, 20)  Then	'������� �������� ����!~200707	
	else
	'==============================================
	'�������� 
	'==============================================
	
			If Not cDataExists.A101.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "��3ȣ ���μ�����ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
			Else
	
				' -- �ڵ�(01)��꼭���������_�����ļ���(W2)
				' ��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(01) ��꼭��������� �� ��ġ 
				
				'call svrmsgbox (UNICDbl(cDataExists.A101.W01, 0) ,0,1)
				'call svrmsgbox (arrVal(2, 1) ,0,1)
				
				If UNICDbl(cDataExists.A101.W01, 0) <> arrVal(2, 1) Then	' -- W1: 01, W2�� �� 
					blnError = True
					Call SaveHTFError(lgsPGM_ID, arrVal(2, 1), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��꼭���������_�����ļ���","��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(01) ��꼭���������"))
				End If

				' -- �ڵ�(02)�ͱݻ���_�����ļ��� 
				If arrVal(3, 5) > 0 Or arrVal(3, 6) > 0 Then
					' -- (05)�غ�� �Ǵ� (06)Ư������ (3)�����Ѽ��� 0 ���� ũ�� �������� 
				Else
					' ��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(02) �ҵ������ݾ�_�ͱݻ��� �� ��ġ 
					If UNICDbl(cDataExists.A101.W02, 0) <> arrVal(2, 2) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, arrVal(2, 2), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ͱݻ���_�����ļ���","��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(02) �ҵ������ݾ�_�ͱݻ���"))
					End If

					' -- �ڵ�(03)�ձݻ���_�����ļ���(W2)
					' ��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(03) �ҵ������ݾ�_�ձݻ��� �� ��ġ 
					If UNICDbl(cDataExists.A101.W03, 0) <> arrVal(2, 3) Then	
						blnError = True
						Call SaveHTFError(lgsPGM_ID, arrVal(2, 3), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ձݻ���_�����ļ���","��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(03) �ҵ������ݾ�_�ձݻ���"))
					End If
					
				End If

				
			End If

			' -- ����� Ŭ���� �޸� ����	: �ڿ� �� ���� 
			Set cDataExists.A101 = Nothing
			' -- �ڵ�(05)�غ��-�����Ѽ�(3) , �ڵ�(06)Ư����_�����Ѽ�(3)  > 0 ��, Ư�������������(A108) �ʼ� �Է� 

			If arrVal(3, 5) > 0 Or arrVal(3, 6) > 0 Then

				Set cDataExists.A108 = new C_TB_5	' -- W4105MA1_HTF.asp �� ���ǵ� 
								
				' -- �߰� ��ȸ������ �о�´�.
				Call SubMakeSQLStatements_W6127MA1("A108",iKey1, iKey2, iKey3)   
								
				cDataExists.A108.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
				cDataExists.A108.WHERE_SQL = lgStrSQL	' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
				'zzzzz �� �� ����?
				If Not cDataExists.A108.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "��5ȣ Ư�������������", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
				End If
				
				Set cDataExists.A108 = Nothing
				
			End If

			' -- 200603 ���� : ����ǥ�������� 0���� Ŭ ��� ����ǥ������ �������(A224) �׸�(7) ����ǥ�����Ͱ� ��ġ (�̰���)
			If arrVal(5, 24) > 0 Then
	
			End If
	
						
			' ��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(12) ���⼼�� �� ��ġ 
			'zzzz Ȯ�ι��� 
			'If UNICDbl(cDataExists.A101.W12, 0) <> arrVal(5, 20) Then	
			'	blnError = True
			'	Call SaveHTFError(lgsPGM_ID, arrVal(5, 20) & " <> " & cDataExists.A101.W12, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(20) ���⼼��","��3ȣ ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(12) ���⼼��"))
			'End If


			' �ڵ�(23) �׸�(5) >= �ڵ�(20) �׸�(3)
			If arrVal(5, 23) >= arrVal(3, 20) Then	
			Else
				blnError = True
				Call SaveHTFError(lgsPGM_ID, arrVal(5, 20), "�ڵ�(23) ��������_�����ļ��� ��(��) �ڵ�(20) ���⼼��_�����Ѽ� ���� ���ų� Ŀ�� �մϴ�.")
			End If

	end if				

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_4 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6127MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A101" '-- �ܺ� ���� SQL

			lgStrSQL = ""

	End Select
	PrintLog "SubMakeSQLStatements_W6127MA1 : " & lgStrSQL
End Sub
%>
