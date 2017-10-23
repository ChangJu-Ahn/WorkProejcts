<%
'======================================================================================================
'*  1. Function Name        : ����1ȣ �������������� 
'*  3. Program ID           : W9123MA1
'*  4. Program Name         : W9123MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_JS1

Set lgcTB_JS1 = Nothing	' -- �ʱ�ȭ 

Class C_TB_JS1
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W2_ETC
	Dim W3
	Dim W3_ETC
	Dim W4
	Dim W4_ETC
	Dim W4_1
	Dim W5
	Dim W5_ETC
	Dim W6
	Dim W6_ETC
	Dim W7_1
	Dim W7_2
	Dim W8
	Dim W9_1
	Dim W9_2
	Dim W9_3
	Dim W9_4
	Dim W9_5
	Dim W9_6
	Dim W9_6__ETC
	Dim W10_1
	Dim W10_2
	Dim W10_3
	Dim W10_4
	Dim W10_5
	Dim W10_6
	Dim W10_7
	Dim W10_8
	Dim W10_9
	Dim W10_9__ETC
		
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

		W1			= oRs1("W1")
		W2			= oRs1("W2")	
		W2_ETC		= oRs1("W2_ETC")
		W3			= oRs1("W3")
		W3_ETC		= oRs1("W3_ETC")
		W4			= oRs1("W4")
		W4_ETC		= oRs1("W4_ETC")
		W4_1		= oRs1("W4_1")
		W5			= oRs1("W5")
		W5_ETC		= oRs1("W5_ETC")
		W6			= oRs1("W6")
		W6_ETC		= oRs1("W6_ETC")
		W7_1		= oRs1("W7_1")
		W7_2		= oRs1("W7_2")
		W8			= oRs1("W8")
		W9_1		= oRs1("W9_1")
		W9_2		= oRs1("W9_2")
		W9_3		= oRs1("W9_3")
		W9_4		= oRs1("W9_4")
		W9_5		= oRs1("W9_5")
		W9_6		= oRs1("W9_6")
		W9_6__ETC	= oRs1("W9_6__ETC")
		W10_1		= oRs1("W10_1")
		W10_2		= oRs1("W10_2")
		W10_3		= oRs1("W10_3")
		W10_4		= oRs1("W10_4")
		W10_5		= oRs1("W10_5")
		W10_6		= oRs1("W10_6")
		W10_7		= oRs1("W10_7")
		W10_8		= oRs1("W10_8")
		W10_9		= oRs1("W10_9")
		W10_9__ETC	= oRs1("W10_9__ETC")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_JS1	A  WITH (NOLOCK) " & vbCrLf	' ����1ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9123MA1
	Dim A100

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9123MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9123MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9123MA1"
	
	Set lgcTB_JS1 = New C_TB_JS1		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_JS1.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 

	'==========================================
	
	
	' -- ����1ȣ �������������� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkBoundary("1,2,3,4", lgcTB_JS1.W1, "ȸ�����α׷�(�ý���)�����Ȳ: " & lgcTB_JS1.W1 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W1, 1)
	' OS(�ü��)
	
	If blnError Then Response.Write "ȸ�����α׷�(�ý���)�����=" & blnError & vbCrLf
	
	
	If lgcTB_JS1.W2 <> "" Then
		If Not ChkBoundary("1,2,3,4,5,6", lgcTB_JS1.W2, "OS(�ü��)" & lgcTB_JS1.W2 & " " ) Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W2, 1)
		if UNIChar(lgcTB_JS1.W2, 1) = "6"  and Trim(lgcTB_JS1.W2_ETC) ="" then 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W2, UNIGetMesg(TYPE_CHK_NULL, "OS(�ü��)�� 6-��Ÿ�̸� OS ��Ÿ�� ",""))
		   
		     blnError = True
		End if     
	End If	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W2_ETC, 50)
	'���α׷� ��� 
	If lgcTB_JS1.W3 <> "" Then
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_JS1.W3, "���α׷� ���: " & lgcTB_JS1.W3 & " " ) Then blnError = True
		
	
		'���α׷� ��� (��Ÿ)
		if UNIChar(lgcTB_JS1.W3, 1) = "8"  and Trim(lgcTB_JS1.W3_ETC) ="" then 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W3, UNIGetMesg(TYPE_CHK_NULL, "���α׷� ��� ������ (6)��Ÿ�̸� OS ���α׷� ��� (��Ÿ) ",""))
		     blnError = True
		End if   
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W3, 1)	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W3_ETC, 50)
	'DBMS 
	
	If lgcTB_JS1.W4 <> "" Then 
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_JS1.W4, "DBMS: " & lgcTB_JS1.W4 & " " ) Then blnError = True
	END IF	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4, 1)
	
	'DBMS (��Ÿ)
	if UNIChar(lgcTB_JS1.W4, 1) = "8"  and Trim(lgcTB_JS1.W4_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W4, UNIGetMesg(TYPE_CHK_NULL, "DBMS  ������ (8)��Ÿ�̸� DBMS (��Ÿ) ",""))
	     blnError = True
	End if   
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4_ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4_1, 30)

	if UNIChar(lgcTB_JS1.W1, 1) = "3"  then 
	
	   If  Trim(lgcTB_JS1.W5) =""  then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W5, UNIGetMesg(TYPE_CHK_NULL, "ȸ�����α׷�(�ý���)�����Ȳ  ������ (3)ERP �̸� ERP",""))
	     blnError = True
     	Else
		   If Not ChkBoundary("1,2,3,4,5", lgcTB_JS1.W5, "ERP: " & lgcTB_JS1.W5 & " " ) Then blnError = True
'		   blnError = True '  2006.03 ���� IF �Ѵ� False �̴�.
		End if   
	End if 
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W5, 1)
	if UNIChar(lgcTB_JS1.W5, 1) = "5"  and Trim(lgcTB_JS1.W5_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W5, UNIGetMesg(TYPE_CHK_NULL, "ERP  ������ (5)��Ÿ�̸� ERP(��Ÿ) ",""))
	     blnError = True
	End if   
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W5_ETC, 50)
	
	if UNIChar(lgcTB_JS1.W1, 1) = "4" then 
	   if    Trim(lgcTB_JS1.W6) ="" then
	         Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W6, UNIGetMesg(TYPE_CHK_NULL, "ȸ�����α׷�(�ý���)�����Ȳ  ������ (4)����� ȸ�����α׷��̸� ����� ȸ�����α׷�",""))
	         blnError = True
	         	
	   Else
	          If Not ChkBoundary("1,2,3,4,5", lgcTB_JS1.W6, "����� ȸ�����α׷�: " & lgcTB_JS1.W6 & " " ) Then blnError = True   
	           
	            	
	   end if      

	End if 
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W6, 1)
	
	if UNIChar(lgcTB_JS1.W6, 1) = "5"  and Trim(lgcTB_JS1.W6_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W6, UNIGetMesg(TYPE_CHK_NULL, "����� ȸ�����α׷�  ������ (5)��Ÿ�̸� ����� ȸ�����α׷�(��Ÿ) ",""))
	     blnError = True
	End if 
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W6_ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W7_1, 50)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W7_2, 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	
	If Not ChkBoundary("1,2", lgcTB_JS1.W8, "���ڻ�ŷ� ����: " & lgcTB_JS1.W8 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W8, 1)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_1, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_2, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_3, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_4, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_5, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_6, 1)
	
	if UNIChar(lgcTB_JS1.W9_6, 1) = "Y" and Trim(lgcTB_JS1.W9_6__ETC) = "" then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W9_6__ETC, UNIGetMesg(TYPE_CHK_NULL, "���ڻ�ŷ������� ��Ÿ�̸� ���ڻ�ŷ���Ÿ�� ",""))
	     blnError = True
	End if     
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_6__ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_1, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_2, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_3, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_4, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_5, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_6, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_7, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_8, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_9, 1)
	if UNIChar(lgcTB_JS1.W10_9, 1) = "Y" and Trim(lgcTB_JS1.W10_9__ETC) = "" then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W10_9__ETC, UNIGetMesg(TYPE_CHK_NULL, "���������ý��������� ��Ÿ�̸� ���������ý��۱�Ÿ�� ",""))
	     blnError = True
	End if  
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_9__ETC, 50)

	sHTFBody = sHTFBody & UNIChar("", 42)	' -- ���� 
	
	' ----------- 
		
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	Else
		Response.End 
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	'Set lgcTB_JS1 = Nothing	' -- �޸�����  <-- W8101MA1_HTF���� ����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9123MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
			
			lgStrSQL = ""
			' -- ������ ���� ����Ÿ ���� üũ 
			
	End Select
	PrintLog "SubMakeSQLStatements_W9123MA1 : " & lgStrSQL
End Sub

%>
