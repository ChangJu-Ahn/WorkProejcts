<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��55ȣ �ҵ��ڷ� ���� 
'*  3. Program ID           : W9113MA1
'*  4. Program Name         : W9113MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_55

Set lgcTB_55 = Nothing ' -- �ʱ�ȭ 

Class C_TB_55
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
			GetData = lgoRs1(pFieldNm)
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
				lgStrSQL = lgStrSQL & " FROM TB_55	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9113MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9113MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum, dblw4_Sum, dblw5_Sum ,dblw4_99, dblw5_99
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9113MA1"

	Set lgcTB_55 = New C_TB_55		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_55.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9113MA1

	'==========================================
	' --��55ȣ �ҵ��ڷ� ���� �������� 
	iSeqNo = 1	: sHTFBody = ""
	dblw4_Sum = 0
	dblw5_Sum = 0
	Do Until lgcTB_55.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		If UNICDbl(lgcTB_55.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			'�����󿩴� ��1��, ��������� ��2��, ��Ÿ�ҵ��� ��3��	
			If Not ChkBoundary("1,2,3", lgcTB_55.GetData("W1") , "�������ڵ�" ) Then  blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W1"), "�ҵ汸��") Then blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W2"), "�������") Then blnError = True
			'If Not ChkDate(lgcTB_55.GetData("W2"),"�ҵ汸���ڵ�" & lgcTB_55.GetData("W1") & "�������") Then  blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W3"), "�ҵ�ͼӿ���") Then blnError = True
			'If Not ChkDate(lgcTB_55.GetData("W3"),"�ҵ汸���ڵ�" & lgcTB_55.GetData("W1") & "�ͼӿ���") Then  blnError = True	
		
			
		    If  Len(UNIRemoveDash(lgcTB_55.GetData("W8"))) <> 10  AND Len(UNIRemoveDash(lgcTB_55.GetData("W8")) ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_55.GetData("W8"), UNIGetMesg(TYPE_CHK_CHARNUM, "�ҵ汸���ڵ�" & lgcTB_55.GetData("W1") & "�ֹε�Ϲ�ȣ","10 �̰ų� 13"))
				blnError = True	
			End If
			  dblw5_Sum = dblw5_Sum  + unicdbl(lgcTB_55.GetData("W5"),0)
			  dblw4_Sum = dblw4_Sum  + unicdbl(lgcTB_55.GetData("W4"),0)
			 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("SEQ_NO"), 6)
			dblw4_99 = Unicdbl(lgcTB_55.GetData("W4"),0)
			dblw5_99 = Unicdbl(lgcTB_55.GetData("W5"),0)
		End If
		
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W1"), 1)
				
		sHTFBody = sHTFBody & UNI8Date(lgcTB_55.GetData("W2"))
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W3"), 4)
	
		If  ChkNotNull(lgcTB_55.GetData("W4"), "���.�� �� ��Ÿ�ҵ�ݾ�") Then 
		   
		Else
			blnError = True
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_55.GetData("W4"), 15, 0)
		
		If  Not ChkNotNull(lgcTB_55.GetData("W5"), "��õ¡���Ҽҵ�ݾ�") Then   blnError = True
		
	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_55.GetData("W5"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W6"), 70)
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W7"), 30)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_55.GetData("W8")), 13)
		sHTFBody = sHTFBody & UNIChar("", 32) & vbCrLf	' -- ���� 
	
		iSeqNo = iSeqNo + 1
		
		lgcTB_55.MoveNext 
	Loop
	
	
	  If dblw4_Sum <> dblw4_99 Then 
						
	   Call SaveHTFError(lgsPGM_ID,dblw4_Sum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���.�� �� ��Ÿ�ҵ�ݾ��հ�","�� ���.�� �� ��Ÿ�ҵ�ݾ� ��"))
	   blnError = True	
	 End if
	
	 If dblw5_Sum <> dblw5_99 Then 
						
	   Call SaveHTFError(lgsPGM_ID,dblw5_Sum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��õ¡���Ҽҵ�ݾ��հ�","�� ��õ¡���Ҽҵ�ݾ� ��"))
	   blnError = True	
	End if
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_55 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9113MA1 : " & lgStrSQL
End Sub
%>
