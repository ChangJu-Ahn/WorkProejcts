<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ������8ȣ�����ŷ����� 
'*  3. Program ID           : W9119MA1
'*  4. Program Name         : W9119MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_KJ8

Set lgcTB_KJ8 = Nothing	' -- �ʱ�ȭ 

Class C_TB_KJ8
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
		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 
	
		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly)  = False Then
			If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If	
			LoadData = False  
			Exit Function
		End If

		
		LoadData = True
	End Function

	Function EOF()
		EOF = lgoRs1.EOF
	End Function

	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	End Function	
		
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
			lgStrSQL = " SELECT  * FROM (" & vbCrLf
            lgStrSQL = lgStrSQL & " SELECT CONVERT(INT, A.W1 ) W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10" & vbCrLf
			lgStrSQL = lgStrSQL & "  From  dbo.ufn_TB_KJ1_HOMETAX_GetRef("& pCode1 &","& pCode2 &","& pCode3 &") A" & vbCrLf
			'lgStrSQL = lgStrSQL & " WHERE A.W1 <>3 "	 & vbCrLf
			lgStrSQL = lgStrSQL & "  Union All " & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT  " & vbCrLf
            lgStrSQL = lgStrSQL & "  CONVERT(INT, A.W1 ) W1 ,  Cast (A.W2 as Varchar(15)), Cast (A.W3 as Varchar(15)), Cast (A.W4 as Varchar(15)), Cast (A.W5 as Varchar(15)), " & vbCrLf
            lgStrSQL = lgStrSQL & "   Cast (A.W6 as Varchar(15)),Cast (A.W7 as Varchar(15)), Cast (A.W8 as Varchar(15)), Cast (A.W9 as Varchar(15))  ,  Cast (A.W10 as Varchar(15)) " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_KJ8 A WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & " ) X "	 & vbCrLf
	
			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

            lgStrSQL = lgStrSQL & " ORDER BY  X.W1 ASC" & vbcrlf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9119MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9119MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows, iSeqNo
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists, sTmp2, sTmp1, sTmp21, sTmp11, sSum
    
  ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9119MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9119MA1"
	
	Set lgcTB_KJ8 = New C_TB_KJ8		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_KJ8.LoadData	Then Exit Function		' -- ��1ȣ ���� �ε� 
		
	'==========================================
	' -- ������8ȣ�����ŷ����� �� �������� 
	
	iSeqNo = 1
    
	For iDx = 2 To 10
		sTmp21 = 0
		sTmp11 = 0
		sTmp2  = 0
		sTmp1  = 0
		sSum   = 0
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' �Ϸù�ȣ 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), "���θ�(��ȣ)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ8.GetData("W" & iDx), 60)
	    stmp = lgcTB_KJ8.GetData("W" & iDx)
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "������(�ּ�)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ8.GetData("W" & iDx), 70)
		
		lgcTB_KJ8.MoveNext 
			
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "�־��� �ڵ�") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ8.GetData("W" & iDx), 7)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "�����ΰ��ǰ���") Then blnError = True	
		
		If Not   ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_KJ8.GetData("W" & iDx), stmp & "�����ΰ��ǰ���")	 Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ8.GetData("W" & iDx), 1)
		
		lgcTB_KJ8.MoveNext 
		
		If  ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), "�հ�") Then 
		    '�հ� =  �׸� (12)����ŷ�_�Ұ� + (13)���԰ŷ�_�Ұ� 
		    sSum = UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)
		    
		Else
			blnError = True		
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx),stmp & "����ŷ�_��") Then blnError = True		
		   sTmp1 = UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx),stmp & "����ŷ�_�����ڻ�") Then blnError = True
		   sTmp11 = UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "����ŷ�_�����ڻ�") Then blnError = True
		   sTmp11 = sTmp11 + UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx),stmp & "����ŷ�_�뿪�ŷ�") Then blnError = True		
		   sTmp11 = sTmp11 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "����ŷ�_�������") Then blnError = True		
		   sTmp11 = sTmp11 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "����ŷ�_��Ÿ") Then blnError = True		
		   sTmp11 = sTmp11 + UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		  
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_�Ұ�") Then blnError = True
		   sTmp2 = UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)		
		   
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_�����ڻ�") Then blnError = True		
		   sTmp21 = UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_�����ڻ�") Then blnError = True		
		   sTmp21 = sTmp21 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_�뿪�ŷ�") Then blnError = True		
		   sTmp21 = sTmp21 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_�������") Then blnError = True		
		   sTmp21 = sTmp21 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ8.MoveNext 
		If Not ChkNotNull(lgcTB_KJ8.GetData("W" & iDx), stmp & "���԰ŷ�_��Ÿ") Then blnError = True		
		   sTmp21 = sTmp21 +  UNICDbl(lgcTB_KJ8.GetData("W" & iDx),0)				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ8.GetData("W" & iDx), 15, 0)
		
		
		
		'�հ� =  �׸� (12)����ŷ�_�Ұ� + (13)���԰ŷ�_�Ұ� 
		if sSum <> sTmp1 + sTmp2 Then
		   Call SaveHTFError(lgsPGM_ID,sSum, UNIGetMesg(TYPE_CHK_NOT_EQUAL,stmp & "�հ�" ,"(12)����ŷ�_�Ұ� + (13)���԰ŷ�_�Ұ�"))
		   
		   blnError = True		
		End if
		'(12)����ŷ�_�Ұ� = [����ŷ�]�����ڻ� + [����ŷ�]�����ڻ� + [����ŷ�]�뿪�ŷ�+ [����ŷ�]������� + [����ŷ�]��Ÿ 
		
		if sTmp1 <> sTmp11 Then
		    Call SaveHTFError(lgsPGM_ID,sTmp1, UNIGetMesg(TYPE_CHK_NOT_EQUAL,stmp  & "(12)����ŷ�_�Ұ�" ,"[����ŷ�]�����ڻ� + [����ŷ�]�����ڻ� + [����ŷ�]�뿪�ŷ�+ [����ŷ�]������� + [����ŷ�]��Ÿ"))
		   
		   blnError = True		
		End if
		
		' (13)���԰ŷ� = [���԰ŷ�]�����ڻ� + [���԰ŷ�]�����ڻ� + [���԰ŷ�]�뿪�ŷ�+ [���԰ŷ�]������� + [���԰ŷ�]��Ÿ 
		
		if sTmp2 <> sTmp21 Then
		    Call SaveHTFError(lgsPGM_ID,sTmp21, UNIGetMesg(TYPE_CHK_NOT_EQUAL,stmp  & "(13)���԰ŷ�_�Ұ�" ,"[���԰ŷ�]�����ڻ� + [���԰ŷ�]�����ڻ� + [���԰ŷ�]�뿪�ŷ�+ [���԰ŷ�]������� + [���԰ŷ�]��Ÿ"))
		   
		  blnError = True		
		End if
		
		
		sHTFBody = sHTFBody & UNIChar("", 55) & vbCrLf	' -- ���� 
	
		lgcTB_KJ8.MoveFirst 

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' �������� �ƴҶ�, ���� ������ ����Ż�� 
			If lgcTB_KJ8.GetData("W" & iDx+1) = "" Then Exit For
		End If
	
	Next
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
'zzzzzz 200703
blnError = false
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_KJ8 = Nothing	' -- �޸�����  
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9119MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- �ܺ� ���� �ݾ� 
	
	End Select
	PrintLog "SubMakeSQLStatements_W9119MA1 : " & lgStrSQL
End Sub

%>
