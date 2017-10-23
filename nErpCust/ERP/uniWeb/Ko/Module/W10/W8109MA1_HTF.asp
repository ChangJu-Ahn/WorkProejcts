<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��13ȣ �����Ư������ 
'*  3. Program ID           : W8109MA1
'*  4. Program Name         : W8109MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_13

Set lgcTB_13 = Nothing ' -- �ʱ�ȭ 

Class C_TB_13
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
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

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
		Call MoveFirst(pType)
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
	
	Function MoveFirst(Byval pType)
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
				lgStrSQL = lgStrSQL & " FROM TB_13_A	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_13_B	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W8109MA1
	Dim A106

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W8109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, sMsg ,dblSum,dblW10,strType,strMsg
    Dim iSeqNo
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8109MA1"

	Set lgcTB_13 = New C_TB_13		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_13.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 


	'==========================================
	' --  ��13ȣ �����Ư������ �������� 
	' -- 1. ����׸��԰ŷ��� 
	
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
    dblSum = 0
    
    
    lgcTB_13.Find 1, "W2_CD='101'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='102'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='103'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='104'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='111'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='112'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='113'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='000'"	' -- (7)����� 
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='121'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='122'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='126'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='127'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='128'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='129'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='125'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='131'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='140'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='132'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='133'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='134'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='135'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='136'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='137'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='138'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='141'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='142'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='139'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W1_CD='10'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    ' -- ���ϻ��� ������ ����� 
    ' ------------------------------------------------------------
    lgcTB_13.Find 1, "W2_CD='101'"
    
	Do Until lgcTB_13.EOF(1) 
		
		

		If lgcTB_13.GetData(1, "W1_CD")  ="07" then
		   dblSum = dblSum + UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
		Elseif  lgcTB_13.GetData(1, "W1_CD")  ="10"    then
			dblW10 = UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
		End if
		
		
       sTmp = lgcTB_13.GetData(1, "W2_CD")
		'�׸� (7) + (121) + (122) + (123) + (124) + (125) + (131) + (132) + (133)+ (134) + (135) + (136) + (137) + (138) + (139)
		Select Case sTmp   '���� ���ϱ� ���ؼ� 
			Case "121" , "122" , "123" , "124" , "125" , "131" , "132" , "133", "134" , "135" , "136" , "137" , "138" , "139"
				  dblSum = dblSum + UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
	    End Select
		
		
		Select Case sTmp
			Case "104", "113", "125", "139"
				'sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
				
				If Not ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_���鼼��") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
				
			Case "121"	,"122" , "131" , "132" , "134" , "135", "136", "138"
						Select Case sTmp
						       Case "121" 
									  strType = "A106_W23"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(23) ��������"
						       Case "122" 
									  strType = "A106_W03"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(03)��������"
						       Case "123" 
									  strType = "A106_W14"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(14)��������"
						       Case "131" 
									  strType = "A106_W31"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(31)��������"
						        Case "132" 
									  strType = "A106_W75"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(75)��������"
						       Case "133" 
									  strType = "A106_W78"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(78)��������"      
						        Case "134" 
									  strType = "A106_W35"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(35)��������"       
						                                   
						       Case "135" 
									  strType = "A106_W36"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(36)��������"    
						                                       
								 Case "136" 
									  strType = "A106_W77"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(77)��������" 
						            
						       Case "138" 
									  strType = "A106_W42"
						              strMsg = " �������鼼�׹��߰����μ����հ�ǥ(��)(A106)�� �ڵ�(42)��������"      
						End Select 
			
			      If  ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_���鼼��") Then 
			      
			 
			          if unicdbl(lgcTB_13.GetData(1, "W4"),0) <> unicdbl(Getdata_TB_3_13ho(strType) ,0) then
			             Call SaveHTFError(lgsPGM_ID, lgcTB_13.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_13.GetData(1, "W2") & "_���鼼��",strMsg))
			             blnError = True	
			          End if   
			      Else
			      
			          blnError = True	
			      End if    
			    	'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
			      
			Case Else
				If Not ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_���鼼��") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
		
		End Select

		Call lgcTB_13.MoveNext(1)	' -- 1�� ���ڵ�� 
	Loop

    if UNICDbl(dblSum,0) <> Unicdbl(dblW10,0)  then
       '���鼼���հ� = �׸� (7) + (121) + (122) + (123) + (124) + (125) + (131) + (132) + (133)+ (134) + (135) + (136) + (137) + (138) + (139)
        Call SaveHTFError(lgsPGM_ID, dblSum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���鼼���հ�","�� �׸���� ��"))
        blnError = True	
    End if

	If Not lgcTB_13.EOF(2) Then
		
		If Not ChkNotNull(lgcTB_13.GetData(2, "W1"), "���μ�����ǥ��") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, Replace("W2_VAL","%","")), "����Ư�����ѹ� ��72�� ����") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W3"), "���⼼��") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W4"), "����ǥ��_�ݾ�") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W6"), "���⼼��") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W7"), "���鼼��") Then blnError = True	
	
	End If

	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W1"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, Replace("W2_VAL","%","")), 5, 2)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W3"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W4"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W6"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W7"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 34) & vbCrLf	' -- ���� 
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	

	Set lgcTB_13 = Nothing	' -- �޸����� 
	
End Function


Function Getdata_TB_3_13ho(byval pType)
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- ��8ȣ �� �������鼼�׸��� 
        Set cDataExists = new TYPE_DATA_EXIST_W8109MA1
		Set cDataExists.A106  = new C_TB_8A	' -- W8107MA1_HTF.asp �� ���ǵ� 
		cDataExists.A106.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A106.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ��							
		' -- �߰� ��ȸ������ �о�´�.
		
		Call SubMakeSQLStatements_W8109MA1(pType,iKey1, iKey2, iKey3)   
	        cDataExists.A106.WHERE_SQL = lgStrSQL
		
		If Not cDataExists.A106.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��8ȣ (��)�������鼼�׼�", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
		
			 dblData = UNICDbl(cDataExists.A106.w4,0)
					
	       	
		End If	
						
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A106 = Nothing
		Set cDataExists = Nothing	' -- �޸����� 
	
		Getdata_TB_3_13ho = unicdbl(dblData,0)


End Function


' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W8109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A106_W23" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '23'" 	 & vbCrLf
	  Case "A106_W03" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '03'" 	 & vbCrLf		
	  Case "A106_W14" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '14'" 	 & vbCrLf		
	  Case "A106_W31" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '31'" 	 & vbCrLf		
	  Case "A106_W75" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '75'" 	 & vbCrLf		
	  Case "A106_W78" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '78'" 	 & vbCrLf	
			
	  Case "A106_W35" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '35'" 	 & vbCrLf		
     
     Case "A106_W36" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '36'" 	 & vbCrLf																						

	  Case "A106_W77" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '77'" 	 & vbCrLf			
			
	  Case "A106_W42" '-- �ܺ� ���� SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '42'" 	 & vbCrLf																						
																					

	End Select
	PrintLog "SubMakeSQLStatements_W8109MA1 : " & lgStrSQL
End Sub
%>
