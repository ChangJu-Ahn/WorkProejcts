
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��1ȣ ���μ�����ǥ�� �Ű� 
'*  3. Program ID           : W8107MA1
'*  4. Program Name         : W8107MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_1

Set lgcTB_1 = Nothing	' -- �ʱ�ȭ 

Class C_TB_1
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W1_RATE
	Dim W1_RATE_View
	Dim W2
	Dim W2_A
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W8
	Dim W8_A
	Dim W9
	Dim W10
	Dim W11
	Dim W12_A
	Dim W12_B
	Dim W13
	Dim W14
	Dim W15
	Dim W16
	Dim W17_1
	Dim W17_2
	Dim W17_Sum
	Dim W18_1
	Dim W18_2
	Dim W18_Sum
	Dim W19_1
	Dim W19_2
	Dim W19_Sum
	Dim W20_1
	Dim W20_2
	Dim W20_Sum
	Dim W21_1
	Dim W21_2
	Dim W21_Sum
	Dim W22_1
	Dim W23_1
	Dim W_TYPE
		
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

		lgStrSQL = ""
		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W1				= oRs1("W1")
		W1_RATE			= oRs1("W1_RATE")
		W1_RATE_VIEW	= oRs1("W1_RATE_VIEW")
		W2				= oRs1("W2")
		W2_A			= oRs1("W2_A")
		W3				= oRs1("W3")
		W4				= oRs1("W4")
		W5				= oRs1("W5")
		W6				= oRs1("W6")
		W7				= oRs1("W7")
		W8				= oRs1("W8")
		W8_A			= oRs1("W8_A")
		W9				= oRs1("W9")
		W10				= oRs1("W10")
		W11				= oRs1("W11")
		W12_A			= oRs1("W12_A")
		W12_B			= oRs1("W12_B")
		W13				= oRs1("W13")
		W14				= oRs1("W14")
		W15				= oRs1("W15")
		W16				= oRs1("W16")
		W17_1			= oRs1("W17_1")
		W17_2			= oRs1("W17_2")
		W17_SUM			= oRs1("W17_SUM")
		W18_1			= oRs1("W18_1")
		W18_2			= oRs1("W18_2")
		W18_SUM			= oRs1("W18_SUM")
		W19_1			= oRs1("W19_1")
		W19_2			= oRs1("W19_2")
		W19_SUM			= oRs1("W19_SUM")
		W20_1			= oRs1("W20_1")
		W20_2			= oRs1("W20_2")
		W20_SUM			= oRs1("W20_SUM")
		W21_1			= oRs1("W21_1")
		W21_2			= oRs1("W21_2")
		W21_SUM			= oRs1("W21_SUM")
		W22_1			= oRs1("W22_1")
		W23_1			= oRs1("W23_1")
		W_TYPE			= oRs1("W_TYPE")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_1	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W8107MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8107MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8107MA1"
	
	Set lgcTB_1 = New C_TB_1		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_1.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ��1ȣ ���μ�����ǥ�ؽŰ� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkBoundary("1,2,3", lgcTB_1.W1, "���α���: " & lgcTB_1.W1 & " " ) Then blnError = True
    sHTFBody = sHTFBody & UNIChar(lgcTB_1.W1, 1)
	If lgcTB_1.W1 = "3" Then
		If UNICDbl(lgcTB_1.W1_RATE, 0) = 0 Then	' �ڵ� 3 �϶� 0���� Ŀ�� �� 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W1, UNIGetMesg(TYPE_CHK_ZERO_OVER, "���� ����",""))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W1_RATE, 6, 3)
	Else
		sHTFBody = sHTFBody & UNINumeric(0, 6, 3)
	End If
	
	' ����ڹ�ȣ(4:2)�� '84'�ε� ���α����� '2'�� �ƴϸ� ����. ('3'�ε� '84'�� ����)
	If (lgcTB_1.W1 <> "2" OR lgcTB_1.W1 <> "3" ) And GetRgstNo42(lgcCompanyInfo.OWN_RGST_NO) = "84"    Then	
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W1, UNIGetMesg("����ڹ�ȣ(4:2)�� '84'�ε� ���α����� '2'�� �ƴϸ� ����. ('3'�ε� '84'�� ����)", "",""))
	End If
	
	If Not ChkBoundary("11,12,21,22,30,40,50,60,70", lgcTB_1.W2, "����������: " & lgcTB_1.W2 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W2, 2)

	If Not ChkBoundary("1,2", lgcTB_1.W3, "��������: " & lgcTB_1.W3 & " " ) Then blnError = True
	If lgcCompanyInfo.EX_RECON_FLG	= "Y" And lgcTB_1.W3 = "2" Then		' ���������� �ܺ������ε�, �ڱ������̶�� üũ�ϸ� ���� 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "��������: �ڱ�", UNIGetMesg("���α������������� �ܺ��������ΰ� '��' �Դϴ�", "",""))
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W3, 1)
	

	
	If Not ChkBoundary("1,2", lgcTB_1.W4, "�ܺΰ��翩��: " & lgcTB_1.W4 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W4, 1)
	
	If Not ChkNotNull(lgcTB_1.W5, "���Ȯ����") Then blnError = True
	If DateDiff("m", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W5) < 0 Then 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W5, UNIGetMesg("����������ں��� �۽��ϴ�.", "",""))
	End If
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W5)
	
	If Not ChkNotNull(lgcTB_1.W6, "�Ű���") Then blnError = True
	If DateDiff("d", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W6) < 0 Or _
	   DateDiff("m", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W6) > 3  Then 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W6, UNIGetMesg("����������ں��� �۰ų�, 3������ �ʰ��Ͽ����ϴ�", "",""))
	ElseIf DateDiff("d", Date(), lgcTB_1.W6 ) < 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W6, UNIGetMesg("�Ű����� ���纸�� �����Դϴ�", "",""))
	End If
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W6)


	sHTFBody = sHTFBody & "10"	' -- �Ű��� 

	' ------------- ����Ÿ ���� üũ�� ���� -------------------
	Set cDataExists	= new TYPE_DATA_EXIST_W8107MA1

	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 
		
	' --- ����Ÿ ��ȸ SQL
	Call SubMakeSQLStatements_W8107MA1("O",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

	If   FncOpenRs("R",lgObjConn,oRs2,lgStrSQL, "", "") = False Then
		blnError = True
	    'Call SaveHTFError(lgsPGM_ID, "�����ļ��Աݾ׸���(A111)", TYPE_DATA_NOT_FOUND)
	    'Call SaveHTFError(lgsPGM_ID, "���ΰ���ǥ�ع׼���������꼭(A101)", TYPE_DATA_NOT_FOUND)
	
	Else
		If oRs2("W_TYPE") = "0" Then
			cDataExists.A130 = "" & oRs2("W_21")	' ���� ���� üũ 0: ����, 1: ���� 
			cDataExists.A131 = "" & oRs2("W_19")
			cDataExists.A132 = "" & oRs2("W_20")
			cDataExists.A170 = "" & oRs2("W_22")
		Else
			cDataExists.A130 = "0" & oRs2("W_21")	
			cDataExists.A131 = "0" & oRs2("W_19")
			cDataExists.A132 = "0" & oRs2("W_20")
			cDataExists.A170 = "0" & oRs2("W_22")
		End If
		oRs2.MoveNext		' W_TYPE ������ ���� ���� ���ڵ�� �ѱ� 
	End If

	'PrintLog "----------.. : " & sHTFBody
	If Not ChkBoundary("1,2", lgcTB_1.W9, "�ֽĺ�������: " & lgcTB_1.W9 & " " ) Then blnError = True
	If lgcTB_1.W9 = "1" Then
		If cDataExists.A131 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W9, UNIGetMesg("�ֽĵ����Ȳ����(A131)�ڷᰡ �����ϴ�", "",""))
		End If
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W9, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W10, "�������ȭ����: " & lgcTB_1.W10 & " " ) Then blnError = True
	If lgcTB_1.W10 = "1" Then
		If cDataExists.A130 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W10, UNIGetMesg("��������������(A130)�ڷᰡ �����ϴ�", "",""))
		End If	
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W10, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W11, "����⵵������: " & lgcTB_1.W11 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W11, 1)

	If Not ChkDate(lgcTB_1.W12_A, "�Ű�Ⱓ������� - ��û") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W12_A)

	If Not ChkDate(lgcTB_1.W12_B, "�Ű�Ⱓ������� - �������") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W12_B)

	If Not ChkBoundary("1,2", lgcTB_1.W13, "��ձݼұް��� ���μ�ȯ�޽�û��: " & lgcTB_1.W13 & " " ) Then blnError = True
	If lgcTB_1.W13 = "1" Then
		If cDataExists.A170 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W13, UNIGetMesg("�ұް������μ���ȯ�޽�û��(A170)�ڷᰡ �����ϴ�", "",""))
		End If	
	End If 
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W13, 1)
 
	If Not ChkBoundary("1,2", lgcTB_1.W14, "�����󰢹���Ű����⿩��: " & lgcTB_1.W14 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W14, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W15, "����ڻ�� �򰡹���Ű� ���⿩��: " & lgcTB_1.W15 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W15, 1)
		
	If UNICDbl(lgcTB_1.W16, 0) < 0 Then	' -- ���Աݾ� ���� ����üũ 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W16, UNIGetMesg(TYPE_CHK_ZERO_OVER, "���Աݾ�",""))
	End If
	
	If oRs2("W_TYPE") = "1" Then
		'-- �����ļ��Աݾ׸���(A111)�� �ڵ�(99)���Աݾ� �հ���  �׸�(4)��� ��ġ���� ������ ���� 
		'--  (�����ļ��Աݾ׸���(A111)���� ���Աݾ��� ��0������ ū ���) 

		sTmp = UNICDbl(oRs2("W_19"), 0)

		If UNICDbl(lgcTB_1.W16, 0) <> sTmp Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W16, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���Աݾ�","�����ļ��Աݾ׸���(A111)�� �ڵ�(99)���Աݾ� �հ��� �׸�(4)��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W16, 15, 0)
			
		oRs2.MoveNext	' -- �������ڵ� 
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "�����ļ��Աݾ׸���(A111)", TYPE_DATA_NOT_FOUND)
	End If
	
	

	If oRs2("W_TYPE") = "2" Then
		' 20�� ����ǥ��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ�ذ� ��ġ 
		' 21�� ����ǥ��_������ �絵�ҵ濡 ���� ���� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(34)����ǥ�ذ� ��ġ 
		' 22�� ����ǥ��_��: ����ǥ��_���μ�(31) + ����ǥ��_������ �絵�ҵ濡 ���� ���μ�(31) 
		If UNICDbl(lgcTB_1.W17_1, 0) <> UNICDbl(oRs2("W_20"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_1 & " <> " & oRs2("W_20") , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ��_���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W17_2, 0) <> UNICDbl(oRs2("W_21"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ��_������ �絵�ҵ濡 ���� ���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(34)����ǥ��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_2, 15, 0)
			
		
		
		If UNICDbl(lgcTB_1.W17_SUM, 0) <> UNICDbl(lgcTB_1.W17_1, 0)  +  UNICDbl(lgcTB_1.W17_2, 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ��_��","����ǥ��_���μ� + ����ǥ��_������ �絵�ҵ濡 ���� ���μ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_SUM, 15, 0)
			
		'23. ���⼼��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(16)���⼼��_�հ�� ��ġ 
		'24. ���⼼��_������ �絵�ҵ濡 ���� ���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(36)���⼼�װ� ��ġ 
		'25. ���⼼��_��	: ���⼼��_���μ�(32)  +  ���⼼��_���� �� �絵�ҵ濡 ���� ���μ�(32)
		If UNICDbl(lgcTB_1.W18_1, 0) <> UNICDbl(oRs2("W_23"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���⼼��_���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(16)���⼼��_�հ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W18_2, 0) <> UNICDbl(oRs2("W_24"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���⼼��_������ �絵�ҵ濡 ���� ���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(36)���⼼��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W18_SUM, 0) <> UNICDbl(lgcTB_1.W18_1, 0)  + UNICDbl(lgcTB_1.W18_2, 0)  Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���⼼��_��", "���⼼��_���μ� + ���⼼��_������ �絵�ҵ濡 ���� ���μ�"))
		End If	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_SUM, 15, 0)
					

		'26. �Ѻδ㼼��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(21)�����Ҽ���_������ + �ڵ�(29)�����Ҽ���_������߰����� 
		'27. �Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�: ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(41)�絵�ҵ���μ�_������� ��ġ 
		'28. �Ѻδ㼼��_�� : �Ѻδ㼼��_���μ�(33) + �Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�(33)			
		If UNICDbl(lgcTB_1.W19_1, 0) <> UNICDbl(oRs2("W_26"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ѻδ㼼��_���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(21)�����Ҽ���_������ + �ڵ�(29)�����Ҽ���_������߰�����"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W19_2, 0) <> UNICDbl(oRs2("W_27"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(41)�絵�ҵ���μ�_������"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W19_SUM, 0) <> UNICDbl(lgcTB_1.W19_1, 0) + UNICDbl(lgcTB_1.W19_2, 0)   Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ѻδ㼼��_��","�Ѻδ㼼��_���μ� + �Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_SUM, 15, 0)
			
	
		'29. �ⳳ��_���μ� : ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(28)���μ� �ⳳ�μ���_�հ�� ��ġ 
		'30. �ⳳ�μ���_������ �絵�ҵ濡 ���� ���μ�: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(44)�絵�ҵ� �ⳳ�μ���_��� ��ġ 
		'31. �ⳳ�μ���_�� : �ⳳ�μ���_���μ�(34) + �ⳳ�μ�_�������� �絵�ҵ濡 ���� ���μ�(34)
		If UNICDbl(lgcTB_1.W20_1, 0) <> UNICDbl(oRs2("W_29"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ⳳ��_���μ�","���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(28)���μ� �ⳳ�μ���_�հ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W20_2, 0) <> UNICDbl(oRs2("W_30"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ⳳ�μ���_������ �絵�ҵ濡 ���� ���μ�","���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(44)�絵�ҵ� �ⳳ�μ���_��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W20_SUM, 0) <> UNICDbl(lgcTB_1.W20_1, 0) + UNICDbl(lgcTB_1.W20_2, 0)   Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ⳳ�μ���_��","�ⳳ��_���μ� + �ⳳ�μ���_������ �絵�ҵ濡 ���� ���μ�"))
		End If		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_SUM, 15, 0)
				
			
		'32. ���������Ҽ���_���μ�	: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(30)���μ� ���������Ҽ��װ� ��ġ 
		'33. ���������Ҽ���_������絵�ҵ濡 ���� ���μ�: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(45)�絵�ҵ� ���������Ҽ��װ� ��ġ 
		'34. ���������Ҽ���_�� : (35)���������Ҽ���_���μ� + (35)���������Ҽ���_������ �絵�ҵ濡 ���� ���μ� 
		If UNICDbl(lgcTB_1.W21_1, 0) <> UNICDbl(oRs2("W_32"), 0) Then
		
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���������Ҽ���_���μ�","���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(28)���μ� �ⳳ�μ���_�հ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W21_2, 0) <> UNICDbl(oRs2("W_33"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���������Ҽ���_������絵�ҵ濡 ���� ���μ�","���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(45)�絵�ҵ� ���������Ҽ���"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W21_SUM, 0) <> UNICDbl(lgcTB_1.W21_1, 0) + UNICDbl(lgcTB_1.W21_2, 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���������Ҽ���_��","���������Ҽ���_���μ� + ���������Ҽ���_������絵�ҵ濡 ���� ���μ�"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_SUM, 15, 0)
								

		'35. �г��Ҽ���	: ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(50)�г��Ҽ���_��� ��ġ 
		'36. �������μ��� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(53)�������μ���_��� ��ġ 
		If UNICDbl(lgcTB_1.W22_1, 0) <> UNICDbl(oRs2("W_35"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W22_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�г��Ҽ���","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(50)�г��Ҽ���_��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W22_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W23_1, 0) <> UNICDbl(oRs2("W_36"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W23_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������μ���","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(53)�������μ���_��"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W23_1, 15, 0)
			
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "���ΰ���ǥ�ع׼���������꼭(A101)", TYPE_DATA_NOT_FOUND)
	End If
	
	' -- �������п��� �ܺ������� �̸� üũ������ ������ ����Ÿ ����/ �ڱ������̸� ����Ÿ�� ���� 
	If lgcTB_1.W3 = "1" Then
		If Not ChkNotNull(lgcCompanyInfo.RECON_BAN_NO, "�ܺ�������_�����ݹ�ȣ") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.RECON_MGT_NO, "�ܺ�������_�����ڰ�����ȣ") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.AGENT_NM, "�ܺ�������_����") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.AGENT_RGST_NO, "�ܺ�������_����ڵ�Ϲ�ȣ") Then blnError = True
	End If
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_BAN_NO), 5)
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_MGT_NO), 6)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_NM, 30)
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.AGENT_RGST_NO), 10)
	
	If Trim(lgcCompanyInfo.BANK_CD) <> "" Then ' �����ڵ尡 ����� �ڵ������ ���¹�ȣ�Է�üũ 
		If Not ChkBoundary("02,03,05,06,07,10,11,12,13,14,15,20,21,23,26,27,31,32,34,35,37,39,71,72,73,74,75,81", lgcCompanyInfo.BANK_CD, "����ȯ�ް���_�����ڵ�: " & lgcCompanyInfo.BANK_CD & " " ) Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.BANK_ACCT_NO, "����ȯ�ް���_���¹�ȣ") Then blnError = True
	End If
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_CD, 2)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_DPST, 20)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_ACCT_NO, 20)
	
	If Not ChkBoundary("Y,N", lgcCompanyInfo.EX_54_FLG, "�ֽĺ����ڷ��ü�����⿩��: " & lgcCompanyInfo.EX_54_FLG & " " ) Then blnError = True
	If lgcCompanyInfo.EX_54_FLG = "Y" Then	' -- Y �� A131, A132 ����Ÿ ����� ���� 
		If cDataExists.A131 = "1" Or cDataExists.A132 = "1" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcCompanyInfo.EX_54_FLG, UNIGetMesg("���α�������_�ֽĺ����ڷ��ü�����⿩�ΰ� '��'�� ��� �ֽĵ����Ȳ����(A131)�ڷᰡ �����ϸ� �����Դϴ�", "",""))
		End If
	End If
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.EX_54_FLG, 1)

	If Not ChkNotNull(lgcTB_1.W_TYPE, "������ ����") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W_TYPE, 3)
	
	sHTFBody = sHTFBody & UNIChar("", 26)	' -- ���� 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.

	If Not blnError Then

		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	'Set lgcTB_1 = Nothing	' -- �޸�����  <-- W8101MA1_HTF���� ����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
			
			lgStrSQL = ""
			' -- ������ ���� ����Ÿ ���� üũ 
			lgStrSQL = lgStrSQL & " SELECT '0' W_TYPE " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_54H WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_19 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_54_BPH WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_20 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_JS1 WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_21 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_68 WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_22 " & vbCrLf
			lgStrSQL = lgStrSQL & "		, 0 W_23, 0 W_24, 0 W_25, 0 W_26, 0 W_27, 0 W_28, 0 W_29" & vbCrLf		
			lgStrSQL = lgStrSQL & "		, 0 W_30, 0 W_31, 0 W_32, 0 W_33, 0 W_34, 0 W_35, 0 W_36" & vbCrLf	
			
			' -- 17ȣ ���� ��			
			lgStrSQL = lgStrSQL & " UNION "		 					 & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT '1' W_TYPE, A.W4 W_19" & vbCrLf
			lgStrSQL = lgStrSQL & "		, 0 W_20, 0 W_21, 0 W_22, 0 W_23, 0 W_24, 0 W_25, 0 W_26, 0 W_27, 0 W_28, 0 W_29" & vbCrLf		
			lgStrSQL = lgStrSQL & "		, 0 W_30, 0 W_31, 0 W_32, 0 W_33, 0 W_34, 0 W_35, 0 W_36" & vbCrLf		
			lgStrSQL = lgStrSQL & " FROM TB_17_D1	A " & vbCrLf	
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.CODE_NO = '99'"		 	 & vbCrLf

			' -- ����Ÿ ������ ���� ȣ�� 
			' 20�� ����ǥ��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ�ذ� ��ġ 
			' 21�� ����ǥ��_������ �絵�ҵ濡 ���� ���� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(34)����ǥ�ذ� ��ġ 
			' 22�� ����ǥ��_��: ����ǥ��_���μ�(31) + ����ǥ��_������ �絵�ҵ濡 ���� ���μ�(31) 
			lgStrSQL = lgStrSQL & " UNION "		 					 & vbCrLf
			'lgStrSQL = lgStrSQL & " SELECT	'2' W_TYPE, 0 W_19, A.W10 W_20, A.W34 W_21, A.W31 + A.W34 W_22" & vbCrLf		
			lgStrSQL = lgStrSQL & " SELECT	'2' W_TYPE, 0 W_19, A.W56 W_20, A.W34 W_21, A.W31 + A.W34 W_22" & vbCrLf		' W10 => W56 : 200603 ���� 
			
			'23. ���⼼��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(16)���⼼��_�հ�� ��ġ 
			'24. ���⼼��_������ �絵�ҵ濡 ���� ���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(36)���⼼�װ� ��ġ 
			'25. ���⼼��_��	: ���⼼��_���μ�(32)  +  ���⼼��_���� �� �絵�ҵ濡 ���� ���μ�(32)
			lgStrSQL = lgStrSQL & "		,	A.W16 W_23, A.W36 W_24, A.W16 + A.W36 W_25" & vbCrLf		

			'26. �Ѻδ㼼��_���μ� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(21)�����Ҽ���_������ + �ڵ�(29)�����Ҽ���_������߰����� 
			'27. �Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�: ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(41)�絵�ҵ���μ�_������� ��ġ 
			'28. �Ѻδ㼼��_�� : �Ѻδ㼼��_���μ�(33) + �Ѻδ㼼��_������ �絵�ҵ濡 ���� ���μ�(33)
			lgStrSQL = lgStrSQL & "		,  A.W21 + A.W29 W_26, A.W41 W_27, A.W21 + A.W29 + A.W41 W_28" & vbCrLf		

			'29. �ⳳ��_���μ� : ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(28)���μ� �ⳳ�μ���_�հ�� ��ġ 
			'30. �ⳳ�μ���_������ �絵�ҵ濡 ���� ���μ�: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(44)�絵�ҵ� �ⳳ�μ���_��� ��ġ 
			'31. �ⳳ�μ���_�� : �ⳳ�μ���_���μ�(34) + �ⳳ�μ�_�������� �絵�ҵ濡 ���� ���μ�(34)
			lgStrSQL = lgStrSQL & "		,  A.W28 W_29, A.W44 W_30, A.W29 + A.W44 W_31" & vbCrLf		
			
			'32. ���������Ҽ���_���μ�	: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(30)���μ� ���������Ҽ��װ� ��ġ 
			'33. ���������Ҽ���_������絵�ҵ濡 ���� ���μ�: ���ΰ���ǥ�ع׼���������꼭(A101)�� �ڵ�(45)�絵�ҵ� ���������Ҽ��װ� ��ġ 
			'34. ���������Ҽ���_�� : (35)���������Ҽ���_���μ� + (35)���������Ҽ���_������ �絵�ҵ濡 ���� ���μ� 
			lgStrSQL = lgStrSQL & "		,  A.W30 W_32, A.W45 W_33, A.W32 + A.W45 W_34" & vbCrLf	
			
			'35. �г��Ҽ���	: ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(50)�г��Ҽ���_��� ��ġ 
			'36. �������μ��� : ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(53)�������μ���_��� ��ġ 
			lgStrSQL = lgStrSQL & "		,  A.W50 W_35, A.W53 W_36 " & vbCrLf	
			lgStrSQL = lgStrSQL & " FROM TB_3	A " & vbCrLf	
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			
	End Select
	PrintLog "SubMakeSQLStatements_W8107MA1 : " & lgStrSQL
End Sub

%>
