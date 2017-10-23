<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��8ȣ �������鼼�װ�꼭(1)
'*  3. Program ID           : W6124MA1
'*  4. Program Name         : W6124MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_8_1

Set lgcTB_8_1 = Nothing	' -- �ʱ�ȭ 

Class C_TB_8_1
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
				lgStrSQL = lgStrSQL & " FROM TB_8_1	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6124MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6124MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6124MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6124MA1"
	
	Set lgcTB_8_1 = New C_TB_8_1		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_8_1.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ��8ȣ �������鼼�װ�꼭(1) �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_1"), "�����������Կ� ���� ���μ�����_�������鼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_1"), 15, 0)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_2"), "���ؼսǼ��װ���_�������鼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_2"), 15, 0)
	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W1_3"), 40)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_3"), "��Ÿ1_�������鼼��") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_3"), 15, 0)
	
	If  ChkNotNull(lgcTB_8_1.GetData("W4_SUM"), "�������鼼��_��") Then 
		if unicdbl(lgcTB_8_1.GetData("W4_SUM"),0) <> unicdbl(lgcTB_8_1.GetData("W4_1"),0) + unicdbl(lgcTB_8_1.GetData("W4_2"),0) + unicdbl(lgcTB_8_1.GetData("W4_3"),0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, " �������鼼��_��","�������鼼���� ����"))
		      blnError = True		
		End if
	Else
	   blnError = True		
	End if
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_1"), "���س���") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W5_1"), 40)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_2"), "���ع߻���") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_8_1.GetData("W5_2"))

	If Not ChkNotNull(lgcTB_8_1.GetData("W5_3"), "������û��") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_8_1.GetData("W5_3"))

	If Not ChkNotNull(lgcTB_8_1.GetData("W5_4_GB"), "����") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W5_4_GB"), 20)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_4"), "���μ�") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W5_4"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 53)	' -- ���� 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_8_1 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6124MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_W6124MA1 : " & lgStrSQL
End Sub

%>
