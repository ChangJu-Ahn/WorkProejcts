<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��Ư ��2ȣ ���׸�����û�� 
'*  3. Program ID           : W6105MA1
'*  4. Program Name         : W6105MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_JT2

Set lgcTB_JT2 = Nothing ' -- �ʱ�ȭ 

Class C_TB_JT2
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
				lgStrSQL = lgStrSQL & " FROM TB_JT2	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6105MA1
	Dim A106

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum , sT1_90Amt , st1_SumTax , st1_Tax
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6105MA1"

	Set lgcTB_JT2 = New C_TB_JT2		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_JT2.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 


	'==========================================
	' -- ��15ȣ �ҵ�ݾ������հ�ǥ �������� 
	iSeqNo = 1	: sHTFBody = ""
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	st1_SumTax = 0
	Do Until lgcTB_JT2.EOF 

        
         
		Select Case Trim(lgcTB_JT2.GetData("W3"))
		
		
		  
		
				
				
		
				
			Case "30"	' -- �հ� 
				If Not ChkNotNull(lgcTB_JT2.GetData("W6"), lgcTB_JT2.GetData("W1") & "_��������") Then blnError = True		
				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2.GetData("W6"), 15, 0)
				st1_Tax =  unicdbl(lgcTB_JT2.GetData("W6"),0)
				
			
			Case "61" , ""	' -- ��Ÿ 
			
				sHTFBody = sHTFBody & UNIChar(lgcTB_JT2.GetData("W1"), 30)	' ����� 
				if  Trim(lgcTB_JT2.GetData("SEQ_NO")) ="117" then ' �ٰŹ� ���� 

				    sHTFBody = sHTFBody & UNIChar(lgcTB_JT2.GetData("W2"), 30)
				End if 
				
				If  ChkNotNull(lgcTB_JT2.GetData("W5"), lgcTB_JT2.GetData("W1") & "_��󼼾�") Then 
				    if  unicdbl(lgcTB_JT2.GetData("W5"),0) <> 0 and    Trim(lgcTB_JT2.GetData("W1")) = "" then
				        Call SaveHTFError(lgsPGM_ID, lgcTB_JT2.GetData("W5"), UNIGetMesg("", " ��󼼾��� 0�� �ƴϸ� ��Ÿ�׸��� �ݵ�� �Է��ؾ� �մϴ�.",""))
				        blnError = True		
				    End if
				Else
				    blnError = True		
				End if    
				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2.GetData("W5"), 15, 0)
				
				If Not ChkNotNull(lgcTB_JT2.GetData("W6"), lgcTB_JT2.GetData("W1") & "_��������") Then blnError = True		
				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2.GetData("W6"), 15, 0)
				st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT2.GetData("W6"),0)
				
				
				
			Case Else
				If Not ChkNotNull(lgcTB_JT2.GetData("W5"), lgcTB_JT2.GetData("W1") & "_��󼼾�") Then blnError = True		
				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2.GetData("W5"), 15, 0)
				
				If Not ChkNotNull(lgcTB_JT2.GetData("W6"), lgcTB_JT2.GetData("W1") & "_��������") Then blnError = True		
				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2.GetData("W6"), 15, 0)
				st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT2.GetData("W6"),0)
		End Select

		
		lgcTB_JT2.MoveNext 
	Loop
	
	
	
	
	sHTFBody = sHTFBody & UNIChar("", 29) & vbCrLf	' -- ���� 
	
	
	if unicdbl(st1_SumTax,0) <> unicdbl(st1_Tax,0)  then
	    Call SaveHTFError(lgsPGM_ID, st1_SumTax, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�׸���� ���鼼����","(118)���鼼��"))
	     blnError = True		
	end if 
	
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	
	Set lgcTB_JT2 = Nothing	' -- �޸����� 
	
End Function


' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "A106_08" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	in ( '08', '09') " 	 & vbCrLf	
	  
	  Case "A106_90" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '90'" 	 & vbCrLf	
      
      Case "A106_04" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '04'" 	 & vbCrLf	    
      
       Case "A106_07" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '07'" 	 & vbCrLf	    
      
      Case "A106_16" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '16'" 	 & vbCrLf	      
        
      Case "A106_97" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '97'" 	 & vbCrLf	                          
            
      Case "A106_98" '-- �ܺ� ���� SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '98'" 	 & vbCrLf	                          
            
            
	End Select
	PrintLog "SubMakeSQLStatements_W6105MA1 : " & lgStrSQL
End Sub
%>
