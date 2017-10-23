<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��48ȣ �ҵ汸�а�꼭 
'*  3. Program ID           : W6101MA1
'*  4. Program Name         : W6101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_48

Set lgcTB_48 = Nothing ' -- �ʱ�ȭ 

Class C_TB_48
	' -- ���̺��� �÷����� 
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	
	
	
	Function Clone(Byref pRs)
		Set pRs = lgoRs2.clone
	End Function



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
		Call SubMakeSQLStatements("H1",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

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
			Case "H1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_48H1	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
			
			Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_48H2	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
			
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6101MA1
	Dim A103

End Class




' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6101MA1()
    Dim iKey1, iKey2, iKey3
    Dim arrHTFBody(1), blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrGamMyun(5), i, iPageNo
    Dim sAmt1,sAmt2, sAmt3
    Dim sRate1,sRate2, sRate3
    Dim dblAmt1
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6101MA1"

	Set lgcTB_48 = New C_TB_48		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_48.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W6101MA1

	'==========================================
	' -- ��48ȣ �ҵ汸�а�꼭 �������� 
	' -- 1. ����������� 
	iSeqNo = 0	
    	Call lgcTB_48.Clone(oRs2)	' ���İ����� �ʿ��� ���� ���ڵ���� ���� 
	' ��������� ó�� ���� 
	Do Until lgcTB_48.EOF(1) 
	
		If UNICDbl(lgcTB_48.GetData(1, "W_TYPE") , 0) < 4 Or (UNICDbl(lgcTB_48.GetData(1, "W_TYPE") , 0) >= 4 and Trim(lgcTB_48.GetData(1, "W_NM"))="") Then
			iPageNo = 0		' ��������ȣ 
		Else
		    iPageNo = 1		' ��������ȣ 
		End If
		arrGamMyun(iSeqNo) = "" & lgcTB_48.GetData(1, "W_NM")
		iSeqNo = iSeqNo + 1
		Call lgcTB_48.MoveNext(1)	' -- 2�� ���ڵ�� 
	Loop

	Do Until lgcTB_48.EOF(2) 
		iSeqNo = 0
		
		For i = 0 To iPageNo		' ������ ��ŭ ����\
		   if i = 0 then
		      sAmt1		= "W4_1"
		      sRate1	= "W5_1"
		      sAmt2		= "W4_2"
		      sRate2	= "W5_2"
			  sAmt3		= "W4_3"
		      sRate3	= "W5_3"
		   Elseif i = 1 then
		      sAmt1		= "W4_4"
		      sRate1	= "W5_4"
		      sAmt2		= "W4_5"
		      sRate2	= "W5_5"
			  sAmt3		= "W4_6"
		      sRate3	= "W5_6"
		   
		   end if   
			arrHTFBody(i) = arrHTFBody(i) & "83"
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(i+1, 2, 0)		' ��������ȣ 
		
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo), 50)	' ������ 
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo+1), 50)	' ������ 
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo+2), 50)	' ������ 
		
			If  ChkNotNull(lgcTB_48.GetData(2, "W1_CD"), "���񱸺��ڵ�") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(UNICDbl(lgcTB_48.GetData(2, "W1_CD"), 0), 2)		' ��� 01, Ȩ�ؽ��� 1�� ���ǵǾ� UNICDbl �� 
		
			If  ChkNotNull(lgcTB_48.GetData(2, "W3"), "�հ�") Then
			  
			    if iPageNo = i Then   '�հ� üũ�̱⶧���� �ѹ��� üũ���ָ� �ȴ�.
						Stmp = UNICDbl(lgcTB_48.GetData(2, "W4_1"),0) + UNICDbl(lgcTB_48.GetData(2, "W4_2"),0) +  UNICDbl(lgcTB_48.GetData(2, "W4_3"),0) _
						       +  UNICDbl(lgcTB_48.GetData(2, "W4_4"),0)+ UNICDbl(lgcTB_48.GetData(2, "W4_5"),0) + UNICDbl(lgcTB_48.GetData(2, "W4_6"),0)+ UNICDbl(lgcTB_48.GetData(2, "W6"),0)
						      
						If Unicdbl(lgcTB_48.GetData(2, "W3"),0) <> Stmp Then 
						
						   Call SaveHTFError(lgsPGM_ID,lgcTB_48.GetData(2, "W3"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�" & lgcTB_48.GetData(2, "W1_CD")  & "�հ�","�� ����� �ݾ��� ��"))
						   blnError = True	
						End if
						
			    
			    
						oRs2.MoveFirst	
						if lgcTB_48.GetData(2, "W1_CD") = "03" Then '(3) ���� ������ : �����(1) - �������(2)
						   oRs2.Find  "W1_CD = '01'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '02'"
						   dblAmt1 = dblAmt1 -  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���� ������","�����(1) - �������(2)"))
							  blnError = True	
							End If	
						End if  
			    
						if lgcTB_48.GetData(2, "W1_CD") = "06" Then '(6) �Ǹź�� ������ �� : ������(4) + �����(5)
						   
						   oRs2.Find  "W1_CD = '04'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '05'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ǹź�� ������ �� "," ������(4) + �����(5)"))
							  blnError = True	
							End If	
						End if   
			    
						if lgcTB_48.GetData(2, "W1_CD") = "07" Then '(7) �������� : ����������(3) -�Ǹź�� ������ ��(6)
						   oRs2.MoveFirst
						   oRs2.Find  "W1_CD = '03'"
						   dblAmt1 =   Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '06'"
						   dblAmt1 = dblAmt1 -   Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������� "," ����������(3) -�Ǹź�� ������ ��(6)"))
							  blnError = True	
							End If	
						End if   
			    
						if lgcTB_48.GetData(2, "W1_CD") = "10" Then '(10) �����ܼ��� �� : ������(8) + �����(9)
						   oRs2.Find  "W1_CD = '08'"
						   dblAmt1 =   Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '09'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, " �����ܼ��� �� ","������(8) + �����(9)"))
							  blnError = True	
							End If	
						End if   
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "13" Then '(13) �����ܺ�� �� : ������(11) + �����(12)
						   oRs2.Find  "W1_CD = '11'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '12'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  �����ܺ�� �� ","������(11) + �����(12)"))
							  blnError = True	
							End If	
						End if    
			    
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "14" Then '(14) ������� : ��������(7) + �����ܼ��� ��(10) - �����ܺ�� ��(13)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '07'"
						    dblAmt1 =   Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '10'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '13'"
						    dblAmt1 = dblAmt1 -  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  ������� ","��������(7) + �����ܼ��� ��(10) - �����ܺ�� ��(13)"))
							  blnError = True	
							End If	
						End if    
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "17" Then '(17) Ư������ �� : ������(15) + �����(16)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '15'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '16'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  Ư������ �� ","������(15) + �����(16)"))
							  blnError = True	
							End If	
						End if  
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "20" Then '(20) Ư���ս� �� : ������(18) + �����(19)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '18'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '19'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  Ư���ս� �� ","������(18) + �����(19)"))
							  blnError = True	
							End If	
						End if      
			    
						  If lgcTB_48.GetData(2, "W1_CD") = "21" Then '(21) �� ��������ҵ� �Ǵ� ���� �� �ҵ� : �������(14) + Ư������ ��(17) - Ư���ս� ��(20)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '14'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '17'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '20'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  �� ��������ҵ� �Ǵ� ���� �� �ҵ� ","�������(14) + Ư������ ��(17) - Ư���ս� ��(20)"))
							  blnError = True	
							End If	
						End if      
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "25" Then '(25) ����ǥ�� : ������⵵�ҵ�(21) - �̿���ձ�(22) - ������ҵ�(23) - �ҵ������(24)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '21'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '22'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '23'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '24'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  (25) ����ǥ��","������⵵�ҵ�(21) - �̿���ձ�(22) - ������ҵ�(23) - �ҵ������(24)"))
							  blnError = True	
							End If	
						End if      
			    End if
			Else
			    blnError = True	
			End If    
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W3"), 15, 0)
		
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt1), "����� �� �ݾ�_1") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt1), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate1), "����� �� ����_1") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate1), 5, 2)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt2), "����� �� �ݾ�_2") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt2), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate2), "����� �� ����_2") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate2), 5, 2)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt3), "����� �� �ݾ�_3") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt3), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate3), "����� �� ����_3") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate3), 5, 2)
		    
		    If iPageNo = 0  then
		        If Not ChkNotNull(lgcTB_48.GetData(2, "W6"), "��Ÿ�� �ݾ�") Then blnError = True	
			     arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W6"), 15, 0)
			    If Not ChkNotNull(lgcTB_48.GetData(2, "W7"), "��Ÿ�� ����") Then blnError = True	
			    arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W7"), 5, 2)
			    
		    Elseif iPageNo = 1 and i = iPageNo then
		        If Not ChkNotNull(lgcTB_48.GetData(2, "W6"), "��Ÿ�� �ݾ�") Then blnError = True	
			     arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W6"), 15, 0)
			    If Not ChkNotNull(lgcTB_48.GetData(2, "W7"), "��Ÿ�� ����") Then blnError = True	
			    arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W7"), 5, 2)
			    
			Else
			    arrHTFBody(i) = arrHTFBody(i) &  UNINumeric("0", 15, 0)
			    arrHTFBody(i) = arrHTFBody(i) &  UNINumeric("0", 5, 2)  
		    End if
		    
		    
			If i = 0 Then
				arrHTFBody(i) = arrHTFBody(i) & UNIChar(lgcTB_48.GetData(2, "DESC1"), 30)
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 15) & vbCrLf	' -- ���� 
			Else
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 30)
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 15) & vbCrLf	' -- ���� 
			End If
			iSeqNo = 3
		
			
		Next
		Call lgcTB_48.MoveNext(2)	' -- 1�� ���ڵ�� 
	Loop

	PrintLog "Write2File : " & arrHTFBody(0) 
	PrintLog "Write2File : " & arrHTFBody(1)
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(arrHTFBody(0))
		Call Write2File(arrHTFBody(1))
	End If
			
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_48 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6101MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W6101MA1 : " & lgStrSQL
End Sub
%>
