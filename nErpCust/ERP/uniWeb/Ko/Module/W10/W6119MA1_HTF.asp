<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��8ȣ �������m����(4)
'*  3. Program ID           : W6119MA1
'*  4. Program Name         : W6119MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_8_4

Set lgcTB_8_4 = Nothing ' -- �ʱ�ȭ 

Class C_TB_8_4
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
				lgStrSQL = lgStrSQL & " FROM TB_8_4H	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_8_4D	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W6119MA1
	Dim A100
	Dim A101

End Class



  Function RtnQueryVal(strField,strFrom,strWhere)
        Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    RtnQueryVal = ""
	    Call CommonQueryRs(strField,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    RtnQueryVal = Replace(lgF0,Chr(11),"")
	    If RtnQueryVal = "X" Or trim(RtnQueryVal) = "" Or IsNull(RtnQueryVal) Then
                     
            ObjectContext.SetAbort
            Call SetErrorStatus
		End If
    End Function
    

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W6119MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, dblW2W4
    Dim iSeqNo
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6119MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6119MA1"
    if Getdata_TB_1("A100")= "3" then
				Set lgcTB_8_4 = New C_TB_8_4		' -- �ش缭�� Ŭ���� 
	
				If Not lgcTB_8_4.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
				' -- ������ ���� ���� 


				'==========================================
				' -- ��8ȣ �������m����(4) �������� 
				' -- 1. ���鼼�װ�� 
				sHTFBody = sHTFBody & "83"
				sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
					
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W1"), "����") Then blnError = True	
				sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(1, "W1"), 20)
					
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W2"), "����ǥ�رݾ�_��������") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W2"), 15, 0)
								
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W3"), "����ǥ�رݾ�_�񰨸�����") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W3"), 15, 0)
					
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W4"), "����ǥ�رݾ�_��") Then
					'����ǥ�رݾ�_��= (2)���������ݾ� + (3)�񰨸����ݾ� 
					'- ���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ�ذ� ��ġ 
					If   unicdbl(lgcTB_8_4.GetData(1, "W4"),0) <> unicdbl(lgcTB_8_4.GetData(1, "W2"),0) + unicdbl(lgcTB_8_4.GetData(1, "W3"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ�رݾ�_��","(2)���������ݾ�+ (3)�񰨸����ݾ�"))
					     blnError = True	
					End if
					
					
					If   unicdbl(lgcTB_8_4.GetData(1, "W4"),0) <> unicdbl(Getdata_TB_3("A101_10"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����ǥ�رݾ�_��","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(10)����ǥ�ذ� ��ġ"))
					     blnError = True	
					End if
					
				Else
				     blnError = True	
				End if     	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W4"), 15, 0)
			
				
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W5"), "���⼼��_��") Then
					'���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(12)���⼼��_��� ��ġ(A153�� ������ݾ��� ��0������ ū ��� �ݵ�� �Է� 
					
					
					If   unicdbl(lgcTB_8_4.GetData(1, "W5"),0) <> unicdbl(Getdata_TB_3("A101_12"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���⼼��_��","���μ�����ǥ�ع׼���������꼭(A101)�� �ڵ�(12)���⼼��_��"))
					     blnError = True	
					End if
					
				Else
				     blnError = True	
				End if     	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W5"), 15, 0)
					
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W6"), "���������ҵ����") Then 
				    '�׸� (2)���������ݾ� / (4)����ǥ�رݾ�_��(��,100% �� �ʰ��ϴ� ���, 100%�� ����)
				  
				    if unicdbl(lgcTB_8_4.GetData(1, "W4"),0)  <> 0 then
				        dblW2W4 =  (unicdbl(lgcTB_8_4.GetData(1, "W2"),0) / unicdbl(lgcTB_8_4.GetData(1, "W4"),0) ) 
				          if unicdbl(dblW2W4,0) * 100  > 100 then
					         dblW2W4 = unicdbl(dblW2W4,0) * 100
						  Else
					         dblW2W4 = unicdbl(dblW2W4,0)
					      end if
				      
				    Else
				        dblW2W4 = 0
				    End if    
				    
				    sTemp = UNICDbl(lgcTB_8_4.GetData(1, "W6"),0)/100   '%������ ������ 
				
					if unicdbl(sTemp,0) <> unicdbl(round(dblW2W4,2),0)  Then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���������ҵ����_��"," (2)���������ݾ� / (4)����ǥ�رݾ�_��"))
					     blnError = True	
					   
					End if   
					
				    
				
				Else
				    blnError = True	
				End if
				
				sHTFBody = sHTFBody & UNINumeric(sTemp , 6, 3)
					
			If Not ChkNotNull(lgcTB_8_4.GetData(1, "W7"), "�������") Then blnError = True	
			   sTemp = UNICDbl(lgcTB_8_4.GetData(1, "W7"),0)/100   '%������ ������ 
			   sHTFBody = sHTFBody & UNINumeric(sTemp, 6, 3)
					
			   If  ChkNotNull(lgcTB_8_4.GetData(1, "W8"), "���鼼��") Then
					    '�׸� (5)���⼼�� x (6)���������ҵ���� x (7)�������(+100�� ~ -100���� �������???)
					    dblW5W6W7 = unicdbl(lgcTB_8_4.GetData(1, "W5"),0) *  (unicdbl(lgcTB_8_4.GetData(1, "W6"),0)*0.01)  *  (unicdbl(lgcTB_8_4.GetData(1, "W7"),0) *0.01)
					  if   unicdbl(lgcTB_8_4.GetData(1, "W8"),0) <= unicdbl(dblW5W6W7,0) + 1000000 and unicdbl(lgcTB_8_4.GetData(1, "W8"),0) >= unicdbl(dblW5W6W7) - 1000000  then
					  Else
					        Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W8"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���鼼��"," (5)���⼼�� x (6)���������ҵ���� x (7)�������"))
					       blnError = True	
					  End if
				
				Else
	
				     blnError = True	
				End If     
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W8"), 15, 0)
	
				sHTFBody = sHTFBody & UNIChar("", 37) & vbCrLf	' -- ���� 

					
				Call lgcTB_8_4.MoveNext(1)	' -- 1�� ���ڵ�� 


				PrintLog "WriteLine2File : " & sHTFBody
				' -- ���Ͽ� ����Ѵ�.
				If Not blnError Then
					Call Write2File(sHTFBody)
				End If
	
				blnError = False : sHTFBody = ""
				' -- 2. �Ϲݰ������ 
				iSeqNo = 1	

				' -- ����Ÿ�� ���߿� �Ѱ����� ��.
				Do Until lgcTB_8_4.EOF(2) 

					If lgcTB_8_4.GetData(2, "W_TYPE") = "1" Then	' �Ϲ� 
						sHTFBody = sHTFBody & "84"
					ElseIf lgcTB_8_4.GetData(2, "W_TYPE") = "2" Then ' �ܱ��� 
						sHTFBody = sHTFBody & "85"
					End If
					sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

					If UNICDbl(lgcTB_8_4.GetData(2, "SEQ_NO"), 0) <> 999999 Then
						sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W9"), "����Ƚ��") Then
						   If Not ChkBoundary("0,1,2,3,4" ,"����Ƚ��: " & lgcTB_8_4.GetData(2, "W9")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W10"), "����") Then
						   If Not ChkBoundary("1,2,3,4,5,6" ,"����: " & lgcTB_8_4.GetData(2, "W10")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
						
					
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W11"), "���� ��� ����") Then blnError = True	
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W12"), "��� ����") Then blnError = True	
					
					Else
						sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "SEQ_NO"), 6)
					End If
							
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W9"), 1)
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W10"), 1)
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W11"))
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W12"))
						
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W13"), "�����ں���_�Ѿ�") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W13"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W14"), "�����ں���_�ܱ����ڰ��ں���") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W14"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W15"), "�����ں���_��������ں���") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W15"), 15, 0)
					
					If Not ChkDate(lgcTB_8_4.GetData(2, "W16_1"), "100% ����Ⱓ From") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W16_1"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W16_2"), "100% ����Ⱓ To") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W16_2"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W17_1"), "50% ����Ⱓ From") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W17_1"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W17_2"), "50% ����Ⱓ To") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W17_2"))
						
				
					If UNICDbl(lgcTB_8_4.GetData(2, "SEQ_NO"), 0) <> 999999 Then
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W18"), "���ػ������������") Then
						  If Not ChkBoundary("1,2,3" ,"���ػ������������: " & lgcTB_8_4.GetData(2, "W18")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
				    End if		
						
							
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W18"), 1)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W19"), "������ܱ����ڰ��ں���") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W19"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W20"), "���ػ���������ں���") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W20"), 15, 0)

					If lgcTB_8_4.GetData(2, "W_TYPE") = "1" Then	' �Ϲ� 
					
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W21"), "�������") Then blnError = True	
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W21"), 6, 3)
						
						sHTFBody = sHTFBody & UNIChar("", 56) & vbCrLf	' -- ���� 
					ElseIf lgcTB_8_4.GetData(2, "W_TYPE") = "2" Then ' �ܱ��� 

						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W21"), "�ܱ������ں���") Then blnError = True	
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W21"), 6, 3)
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W35"), "�������") Then 
						   '(�׸�(32)������ܱ����ڰ��ں��� / �׸�(33)�Ѱ�����ܱ����ڰ��ں��� ) x �׸�(34)�ܱ������ں��� 
						    if unicdbl(lgcTB_8_4.GetData(1, "W20"),0) <> 0 then
						       if unicdbl(lgcTB_8_4.GetData(2, "W35"),0) <> (unicdbl(lgcTB_8_4.GetData(2, "W19"),0) /unicdbl(lgcTB_8_4.GetData(2, "W20"),0)) *unicdbl(lgcTB_8_4.GetData(1, "W21"),0) then
						           Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W35"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������ܱ����ڰ��ں���_��","�Ѱ�����ܱ����ڰ��ں��� ) x �ܱ������ں���"))
						           blnError = True	
						       End if 
						    End if   
						Else
						    blnError = True	
						End if    
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W35"), 6, 3)
						
						sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- ���� 
					End If

					iSeqNo = iSeqNo + 1
					
					Call lgcTB_8_4.MoveNext(2)	' -- 2�� ���ڵ�� 
				Loop

				' ----------- 
				'Call SubCloseRs(oRs2)
	
				PrintLog "Write2File : " & sHTFBody
				' -- ���Ͽ� ����Ѵ�.
				If Not blnError Then
					Call Write2File(sHTFBody)
				End If

				
				Set lgcTB_8_4 = Nothing	' -- �޸����� 
		End if
End Function



Function Getdata_TB_1(byval strType)
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- ��8ȣ �� �������鼼�׸��� 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A100  = new C_TB_1	' -- W8107MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W6119MA1(strType,iKey1, iKey2, iKey3)   
				
		
				
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��1ȣ ���μ�����ǥ�ع׼���������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else

			   dblData = UNICDbl(cDataExists.A100.W1, 0)

		End If	
						
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A100 = Nothing
		Set cDataExists = Nothing	' -- �޸����� 
	
		Getdata_TB_1 = unicdbl(dblData,0)


End Function


Function Getdata_TB_3(byval strType )
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- ��8ȣ �� �������鼼�׸��� 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A101  = new C_TB_3	' -- W8101MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
		Call SubMakeSQLStatements_W6119MA1(strType,iKey1, iKey2, iKey3)   
						

						
		If Not cDataExists.A101.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��3ȣ ���μ� ����ǥ�� �� ����������꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
             Select Case  strType
               Case "A101_10"
			        'dblData = UNICDbl(cDataExists.A101.W10, 0)
			        dblData = UNICDbl(cDataExists.A101.W56, 0)	' 2006-01-05 (200603 ������)
			   Case "A101_12"
			        dblData = UNICDbl(cDataExists.A101.W12, 0)
			 End Select  
		
		End If	
						
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A101 = Nothing
		Set cDataExists = Nothing	' -- �޸����� 
	
		Getdata_TB_3 = unicdbl(dblData,0)


End Function




' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W6119MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A100" '-- �ܺ� ���� SQL
	       lgStrSQL =""
      Case "A101" '-- �ܺ� ���� SQL
           lgStrSQL=""
	End Select
	PrintLog "SubMakeSQLStatements_W6119MA1 : " & lgStrSQL
End Sub
%>
