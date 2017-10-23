<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��51ȣ �߼ұ�����ذ���ǥ 
'*  3. Program ID           : WB107MA1
'*  4. Program Name         : WB107MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_51

Set lgcTB_51 = Nothing	' -- �ʱ�ȭ 

Class C_TB_51
	' -- ���̺��� �÷����� 
	
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
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		
		LoadData = True
	End Function

	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_51	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_WB107MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_WB107MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_WB107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "WB107MA1"
	
	Set lgcTB_51 = New C_TB_51		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_51.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ��51ȣ �߼ұ�����ذ���ǥ �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 


    If unicdbl(lgcTB_51.GetData("W07"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W01"), "������Աݾ�1�� 0�� �ƴϸ� ����1") Then blnError = True		
	End if	
		sHTFBody = sHTFBody & UNIChar(Trim(lgcTB_51.GetData("W01")), 30)

	If unicdbl(lgcTB_51.GetData("W07"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W04"), "������Աݾ�1�� 0�� �ƴϸ� ���ذ�����ڵ�1") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W04"), 7)
	
	If Not ChkNotNull(lgcTB_51.GetData("W07"), "������Աݾ�1") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W07"), 15, 0)
	
	If unicdbl(lgcTB_51.GetData("W08"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W02"), "������Աݾ�2�� 0�� �ƴϸ� ����2") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W02"), 30)
	

	
	If unicdbl(lgcTB_51.GetData("W08"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W05"), "������Աݾ�2�� 0�� �ƴϸ� ���ذ�����ڵ�2") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W05"), 7)                 '���ذ�����ڵ�2
	
		
	If Not ChkNotNull(lgcTB_51.GetData("W08"), "������Աݾ�2") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W08"), 15, 0)          ' ������Աݾ�2
	
	If unicdbl(lgcTB_51.GetData("W09"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W06"), "������Աݾ�-��Ÿ�� 0�� �ƴϸ� ���ذ�����ڵ�_��Ÿ") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W06"), 7)                 '���ذ�����ڵ�-��Ÿ 
	
	If Not ChkNotNull(lgcTB_51.GetData("W09"), "������Աݾ�-��Ÿ") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W09"), 15, 0)          ' ������Աݾ�-��Ÿ 

	
	
	If ChkNotNull(lgcTB_51.GetData("W_SUM"), "��_������Աݾ�") Then 
	   '��-������Աݾ� : �׸� (7) + (8) + (9)
	   If UNICDbl(lgcTB_51.GetData("W_SUM"),0)  <> unicdbl(lgcTB_51.GetData("W07"),0) + UniCdbl(lgcTB_51.GetData("W08"),0)+ UniCdbl(lgcTB_51.GetData("W09"),0) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��_������Աݾ�","������Աݾ���"))
	        blnError = True		
	   End if
	Else
		'blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W_SUM"), 15, 0)
	
	If Not ChkBoundary("1,2", lgcTB_51.GetData("W19"), "�ش���_���տ���" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W19"), 1)

	If Not ChkNotNull(lgcTB_51.GetData("W10"), "�����������") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W10"), 4, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W11"), "�߼ұ���⺻������� ��ǥ1�� �Ը����_��������") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W11"), 4, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W12"), "�ں���") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W12"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W13"), "�߼ұ���⺻������� ��ǥ1�� �Ը����_�ں���") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W13"), 6, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W14"), "�����") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W14"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W15"), "�߼ұ���⺻������� ��ǥ1�� �Ը����_�����") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W15"), 6, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W16"), "�ڱ��ں�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W16"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W17"), "����.��ȸ��Ϲ����� ��� �ڻ��Ѿ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W17"), 7, 1)
	
	If Not ChkBoundary("1,2", lgcTB_51.GetData("W20"), "�Ը�_���տ���" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	If Not ChkBoundary("1,2", lgcTB_51.GetData("W21"), "�濵_���տ���" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W21"), 1)


    If lgcTB_51.GetData("W18") <> "" Then 
		If  unicdbl(lgcTB_51.GetData("W18"),0 ) <> 0 and unicdbl(lgcTB_51.GetData("W18"),0 ) <= 2001 Then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_CHK_LOW_AMT, "�ʰ��⵵","2001�̰ų� 2001"))
		     blnError = True
		End if
	Else
		'blnError = True		
	End if		
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W18"), 4, 0)  '�ʰ����� 


	If Not ChkBoundary("1,2", lgcTB_51.GetData("W22"), "�����Ⱓ_���տ���" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	If  ChkBoundary("1,2", lgcTB_51.GetData("W23"), "��������" ) Then
	'- �׸� (19), (20), (21), (22) �� ��� 1(����) �� ��쿡 ���ؼ�  '1' (����) 
	'- �׸�(20)�� '2'(������) �̰� �׸�(22)�� 1(����)�� ��쵵 '1' (����)
	   if lgcTB_51.GetData("W23") <> "1" and lgcTB_51.GetData("W19") ="1" and  lgcTB_51.GetData("W20") = "1" and  lgcTB_51.GetData("W21") = "1" and  lgcTB_51.GetData("W22") = "1" then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_MSG_NORMAL_ERR, "��������",""))
	        blnError = True
	   
	   End if
	   
	   
	     if lgcTB_51.GetData("W23") <> "1" and lgcTB_51.GetData("W20") ="2" and  lgcTB_51.GetData("W22") = "1"  then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_MSG_NORMAL_ERR, "��������",""))
	        
	        blnError = True
	   
	     End if
	  
	   
	Else
	  '  blnError = True
	End if
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	sHTFBody = sHTFBody & UNIChar("", 46)	' -- ���� 

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.

'zzzz  ??

blnError = false

	If Not blnError Then
		
		Call WriteLine2File(sHTFBody)
	End If
 
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_51 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_WB107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_WB107MA1 : " & lgStrSQL
End Sub

%>
