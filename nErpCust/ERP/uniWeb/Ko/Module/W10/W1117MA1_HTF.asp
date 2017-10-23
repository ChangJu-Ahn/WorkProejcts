<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��3ȣ��3(3) �μӸ��� ��Ÿ���� 
'*  3. Program ID           : W1117MA1
'*  4. Program Name         : W1117MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
' -- �� ������ Ŭ������ W1107MA1_HTF �� ����Ȱ� ����Ѵ�.

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W1117MA1
	Dim A115

End Class

Function Clone(Byref pRs)
	Set pRs = lgoRs1.clone
End Function
' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W1117MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim dblAmt1, dblAmt2, dblAmt3, arrNew(50)
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1117MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1117MA1"

	Set lgcTB_3_3_3 = New C_TB_3_3_3		' -- �ش缭�� Ŭ���� 
	
	lgcTB_3_3_3.WHERE_SQL = "		AND A.W1 = '8' "		' 
 
	If Not lgcTB_3_3_3.LoadData Then Exit Function			
	
	
	'==========================================
	' -- ��3ȣ��3(3) �μӸ��� ��Ÿ���� ���ڽŰ� �� �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	Call lgcTB_3_3_3.Clone(oRs2)	' ���İ����� �ʿ��� ���� ���ڵ���� ���� 
	Do Until lgcTB_3_3_3.EOF 
	
		If  ChkNotNull(lgcTB_3_3_3.W5, lgcTB_3_3_3.W3) Then 
	    

					If lgcTB_3_3_3.W4 = "01" Then   '- �ڵ�(01)����= �ڵ� 02 + 03 - 04
						oRs2.Find "W4 = '02'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						oRs2.Find "W4 = '03'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
					
						oRs2.Find "W4 = '04'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 - dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����","�ڵ� 02 + 03 - 04"))
						   blnError = True	
						End If

					End If
		
		
					If lgcTB_3_3_3.W4 = "05" Then	 '- �ڵ�(05)�빫��= �ڵ� 06 + 07 + 08
						oRs2.Find "W4 = '06'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '07'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '08'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 + dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�빫��","�ڵ� 06 + 07 + 08"))
						   blnError = True	
						End If

					End If
		 
		
		
		
		
		
					If lgcTB_3_3_3.W4 = "09" Then   ' - �ڵ�(09)���= �ڵ� 10 + 11 + 12 + 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 22+ 23 + 24 + 25 + 26
						oRs2.Find "W4 = '10'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '11'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '12'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '13'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '14'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '15'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '16'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '17'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '18'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '19'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '20'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '21'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '22'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '23'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '24'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '25'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '26'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						' -- 2006.03 ���� 
						oRs2.MoveFirst
						oRs2.Find "W4 = '33'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "���","�ڵ� 10 + 11 + 12 + 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 22+ 23 + 24 + 25 + 26 + 27 + 28"))
						   blnError = True	
						End If	
						
			
					End If
		
		
					If lgcTB_3_3_3.W4 = "27" Then	 '- �ڵ�(27)����Ѻ��= �ڵ� 01 + 05 + 09
						oRs2.MoveFirst				' ���� ���ڵ�� ���ư��� �˻��ϰ����Ҷ� MoveFirst�ؾߵ� 
						oRs2.Find "W4 = '01'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '05'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '09'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 + dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����Ѻ��","�ڵ� 01 + 05 + 09"))
							blnError = True	
						End If

					End If

				
					If lgcTB_3_3_3.W4 = "29" Then	 '�ڵ�(29)�հ�= �ڵ� 27 + 28
						oRs2.Find "W4 = '27'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '28'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2  Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�հ�"," �ڵ� 27 + 28"))
							blnError = True	
						End If
					End If
		
		
					If lgcTB_3_3_3.W4 = "32" Then	 '�ڵ�(32)���� = �ڵ� 29 - 30 - 31
						oRs2.Find "W4 = '29'"		' �ش��ڵ�� �ݵ�� �����ؾ�, �����࿡�� ������ �ȳ� 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '30'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '31'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)

						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 - dblAmt2 - dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����"," �ڵ� 29 - 30 - 31"))
							blnError = True	
						End If

					End If
						
		Else
		       blnError = True	
		End if
		
		' -- 2006.03 ���� 
		Select Case lgcTB_3_3_3.W4
			Case "33"
				arrNew(33) = lgcTB_3_3_3.W5
			Case Else
				sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_3.W5, 15, 0)
		End Select
		
		lgcTB_3_3_3.MoveNext 

	Loop
	
	' -- 2006.03 �������� 
	sHTFBody = sHTFBody & UNINumeric(arrNew(33), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 14)	' -- ���� 
	
	' ----------- 
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3_3_3 = Nothing	' -- �޸����� 
	
End Function


' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W1117MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W1117MA1 : " & lgStrSQL
End Sub
%>