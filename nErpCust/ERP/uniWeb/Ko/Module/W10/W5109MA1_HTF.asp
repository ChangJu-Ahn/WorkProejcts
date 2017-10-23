<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��22ȣ ��αݸ��� 
'*  3. Program ID           : W5109MA1
'*  4. Program Name         : W5109MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_22

Set lgcTB_22 = Nothing ' -- �ʱ�ȭ 

Class C_TB_22
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.

	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

		' --������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		If blnData1 = False And blnData2 = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 1
				 lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				 lgoRs1.MoveNext
			Case 2
				 lgoRs2.MoveNext
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
				End If
		End Select
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
		Call SubCloseRs(lgoRs2)		
	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_22H	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_22D	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W5109MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W5109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W5109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W5109MA1"

	Set lgcTB_22 = New C_TB_22		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_22.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W5109MA1

	'==========================================
	' -- ��22ȣ ��αݸ��� �������� 
	' -- 1. ����׸��԰ŷ��� 
	'sHTFBody = "83"
	'sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	sHTFBody =""
	iSeqNo = 1	
	
	Do Until lgcTB_22.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		
		If Not ChkNotNull(lgcTB_22.GetData(1, "W2"), lgcTB_22.GetData(1, "W1") & "�����ڵ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W2"), 2)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W3"), 40)
		
		If lgcTB_22.GetData(1, "W2") <> "20" And lgcTB_22.GetData(1, "W2") <> "50" Then	
		    '��¥ ����üũ 
			If Not (ChkNotNull(lgcTB_22.GetData(1, "W4"), lgcTB_22.GetData(1, "W1") & "����") and  ChkDate(lgcTB_22.GetData(1, "W4"),lgcTB_22.GetData(1, "W1") & "����"))  Then blnError = True	
			 
			'������ڰ� ����⵵ �Ⱓ �̳��� �ƴϸ� ���� 
			If Not (UNI6Date(lgcTB_22.GetData(1, "W4"))  >= UNI6Date(lgcCompanyInfo.FISC_START_DT) and UNI6Date(lgcTB_22.GetData(1, "W4"))  <=UNI6Date(lgcCompanyInfo.FISC_END_DT))   Then 
			  Call SaveHTFError(lgsPGM_ID,lgcTB_22.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_22.GetData(1, "W1") & "�������", "����⵵�Ⱓ"))
			  blnError = True	
			End if  
			
		End If
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_22.GetData(1, "W4")), 6)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W5"), 80)
		
		
		If lgcTB_22.GetData(1, "W2") <> "20" Then	
			If Not ChkNotNull(lgcTB_22.GetData(1, "W6"), "���ó") Then blnError = True	
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W6"), 60)
		
		
	   if  	UNIRemoveDash(lgcTB_22.GetData(1, "W7")) <> "" then
			If  Len(UNIRemoveDash(lgcTB_22.GetData(1, "W7"))) <> 10  and Len(UNIRemoveDash(lgcTB_22.GetData(1, "W7")) ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W7"), UNIGetMesg(TYPE_CHK_CHARNUM, lgcTB_22.GetData(1, "W6") & "����ڵ�Ϲ�ȣ(�ֹε�Ϲ�ȣ)","10 �̰ų� 13"))
				blnError = True	
			End If
			   
			If UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO) = UNIRemoveDash(lgcTB_22.GetData(1, "W7"))  Then
			    Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W7"), UNIGetMesg("������ ����ڹ�ȣ�� �ŷ����� ����ڹ�ȣ�� ���� �� �����ϴ�", "",""))
				blnError = True	
			 End If
		End if	 
		 
		 sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_22.GetData(1, "W7")), 13)
		
		If  ChkNotNull(lgcTB_22.GetData(1, "W8"), "�ݾ�") Then 
		    If Unicdbl(lgcTB_22.GetData(1, "W8"),0) < 0 then 
		         Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W8"), UNIGetMesg(TYPE_POSITIVE, "�ݾ�",""))
				blnError = True	
		    End if
		Else
			blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_22.GetData(1, "W8"), 15, 0)
	 
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W_DESC"), 80)
		
		sHTFBody = sHTFBody & UNIChar("", 42) & vbCrLf	' -- ���� 
		
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_22.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop
	
	
	         '999990 : ��. �� ��24����2�� ��α�(������α�, �ڵ� 10)�� ���� 
             '999991 : ��. ����Ư�����ѹ� ��76�� ��α�(��ġ�ڱ�, �ڵ� 20)�� ���� 
             '999992 : ��. ����Ư�����ѹ� ��73�� ��1�� ��1ȣ ��α�(�ڵ� 60)�� ���� 
             '999993 : ��. ����Ư�����ѹ� ��73�� ��1�� ��2ȣ ���� ��15ȣ ��α�(�ڵ� 30)�� ���� 
             '999994 : ��. �� ��24�� ��1�� ��α�(������α�, �ڵ� 40)�� ���� 
             '999995 : ��. ����Ư�����ѹ� ��73�� ��2�� ��α�(�ڵ� 70)�� ���� 
             '999996 : ��. ��Ÿ ��α�(�ڵ� 50)�� ���� 
             '999999 : �հ�μ� 999990 ~ 999996 �� ������ ���մϴ�.

	
	' -- 2. �ں��ŷ� 
	iSeqNo = 999990	
	
	Do Until lgcTB_22.EOF(2) 
		If lgcTB_22.GetData(2, "W9_CD") <> "20" Then	' -- 20 ��ġ�� ���ŵ� 
	
			if   UNIChar(lgcTB_22.GetData(2, "W9_CD"), 2)  = "99" Then 
				iSeqNo = 999999
			End if
	
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 
			If Not ChkNotNull(lgcTB_22.GetData(2, "W9_CD"), "�����ڵ�") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(2, "W9_CD"), 2)
		
			sHTFBody = sHTFBody & UNIChar("", 40)
			sHTFBody = sHTFBody & UNIChar("", 6)
			sHTFBody = sHTFBody & UNIChar("", 80)
			sHTFBody = sHTFBody & UNIChar("", 60)
			sHTFBody = sHTFBody & UNIChar("", 13)
		
			If Not ChkNotNull(lgcTB_22.GetData(2, "W9_AMT"), "�ݾ�") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_22.GetData(2, "W9_AMT"), 15, 0)
		
			sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(2, "W9_DESC"), 80)

			sHTFBody = sHTFBody & UNIChar("", 42) & vbCrLf	' -- ���� 
		End If
				
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_22.MoveNext(2)	' -- 1�� ���ڵ�� 
	Loop
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_22 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W5109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W5109MA1 : " & lgStrSQL
End Sub
%>
