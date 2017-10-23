<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��54ȣ �ֽĵ� ������Ȳ����(��)
'*  3. Program ID           : W9109MA1
'*  4. Program Name         : W9109MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_54

Set lgcTB_54 = Nothing ' -- �ʱ�ȭ 

Class C_TB_54
	' -- ���̺��� �÷����� 
	Dim TAX_OFFICE
	Dim INCOM_DT
	Dim W5
	Dim W6
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs3		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	
	' ------------------ Ŭ���� ����Ÿ �ε� �Լ� --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3
				 
		'On Error Resume Next                                                             '��: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True : blnData3 = True 
		
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
		Call SubMakeSQLStatements("D1",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		' --������ �о�´�.
		Call SubMakeSQLStatements("D2",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
				
		If blnData1 = False And blnData2 = False And blnData3 = False Then
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
			Case 2
				lgoRs2.Find pWhereSQL
			Case 3
				lgoRs3.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 2
				lgoRs2.Filter = pWhereSQL
			Case 3
				lgoRs3.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
			Case 3
				EOF = lgoRs3.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 2
			    If Not lgoRs1.EOF Then
				   lgoRs2.MoveFirst
				End if   
			Case 3
				lgoRs3.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 2
			    If Not lgoRs1.EOF Then
				   lgoRs2.MoveNext
				End if   
			Case 3
				lgoRs3.MoveNext
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
			Case 3
				If Not lgoRs3.EOF Then
					GetData = lgoRs3(pFieldNm)
				End If
		End Select
	End Function
	
	Function RecordCount(Byval pType)
		Select Case pType
			Case 3
				RecordCount = lgoRs3.RecordCount
		End Select
	End Function

	Function CloseRs()	' -- �ܺο��� �ݱ� 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
		
	End Function
		
	'----------- Ŭ���� ����/���� �̺�Ʈ -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- ���ڵ���� ����(����)�̹Ƿ� Ŭ���� �ı��ÿ� �����Ѵ�.
		Call SubCloseRs(lgoRs2)	
		Call SubCloseRs(lgoRs3)

	End Sub

	' ------------------ ��ȸ SQL �Լ� --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
	            lgStrSQL = lgStrSQL & " B.TAX_OFFICE, B.INCOM_DT, B.SUBMIT_FLG  " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.RECON_MGT_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE ''" & vbCrLf
	            lgStrSQL = lgStrSQL & "   END RECON_MGT_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.OWN_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "   END AGENT_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_NM" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.CO_NM " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END AGENT_NM "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN C.W_NAME" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.REPRE_NM " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END REPRE_NM "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN C.W_CO_ADDR" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.CO_ADDR " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END CO_ADDR "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_TEL_NO" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.TEL_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END TEL_NO "& vbCrLf
	            lgStrSQL = lgStrSQL & " , B.OWN_RGST_NO, B.CO_NM, B.REPRE_NM ,HOME_TAX_USR_ID" & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54H	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & "		INNER JOIN TB_COMPANY_HISTORY	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_AGENT_INFO C  WITH (NOLOCK) ON B.CO_CD=C.CO_CD AND B.FISC_YEAR=C.FISC_YEAR AND B.REP_TYPE=C.REP_TYPE AND C.W_TYPE='��ǥ�̻�' " & vbCrLf	' ����3ȣ 

				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54D1	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54D2	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
							
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9109MA1
	Dim A101

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, arrTmp1(1), arrTmp2(1)
    Dim iSeqNo, iDx, iDtlCnt,dblNum1,dblAmtRate1,dblNum2
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9109MA1"

	Set lgcTB_54 = New C_TB_54		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_54.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9109MA1

	'==========================================
	' -- ��54ȣ �ֽĵ� ������Ȳ����(��) �������� 
	' -- 1. �⺻����(83A131)
	sHTFBody = "A"
	sHTFBody = sHTFBody & UNIChar("40", 2)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	
	
	  ' -- ��8ȣ �� �������鼼�׸��� 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A100  = new C_TB_1	' -- W8101MA1_HTF.asp �� ���ǵ� 
							
		
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��1ȣ ���μ�����ǥ�ع׼��׽Ű�(A100)", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
            '- �Ϲݹ����� �񿵸� ����(���α��� : 50, 60, 70)�� ��� �ԷµǸ� ����- ���μ�����ǥ�ع׼��׽Ű�(A100)�� �ֽĺ����ڷ��ü�����⿩�ΰ� 'Y' �� ��� �ԷµǸ� ���� 
		     If cDataExists.A100.W2 = "50" Or cDataExists.A100.W2 = "60" Or cDataExists.A100.W2="70" Then
		        blnError = True	
		        Call SaveHTFError(lgsPGM_ID, cDataExists.A100.W2, UNIGetMesg("�Ϲݹ����� �񿵸� ����(���α��� : 50, 60, 70)�� ��� �ԷµǸ� �ȵ˴ϴ�", "",""))
		        Exit Function
		     End IF
		    
		End If	
						
		' -- ����� Ŭ���� �޸� ���� 
		Set cDataExists.A100 = Nothing
		Set cDataExists = Nothing	' -- �޸����� 
	
		 
		 If lgcCompanyInfo.EX_54_FLG = "Y"  Then   ' �ֽĸ�ü���� 
		   blnError = True	
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg("���μ�����ǥ�ع׼��׽Ű�(A100)�� �ֽĺ����ڷ��ü�����⿩�ΰ� 'Y' �� ��� �ԷµǸ� �ȵ˴ϴ�", "",""))
		   Exit Function
		End IF
		
		
	
	If  ChkNotNull(lgcTB_54.GetData(1, "TAX_OFFICE"), "�������ڵ�") Then 
	
	    If Not ChkNumeric(lgcTB_54.GetData(1, "TAX_OFFICE"),"�������ڵ�") Then
	      blnError = True	
	    End If  
	   
	Else
	    blnError = True	
	End If   
	 sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)
		
	If  ChkNotNull(lgcTB_54.GetData(1, "INCOM_DT"), "��������") Then
	    If Not ChkDate(lgcTB_54.GetData(1, "INCOM_DT"),"��������") Then
	        blnError = True	
	    End If
	Else
	     blnError = True	    
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "INCOM_DT"))
	
		
	If  ChkNotNull(lgcTB_54.GetData(1, "SUBMIT_FLG"), "�����ڱ���") Then
	    '- ��1�� : �����븮��, ��2�� : ������ 
	    If Not ChkBoundary("1,2" , lgcTB_54.GetData(1, "SUBMIT_FLG") , "�����ڱ���") Then
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End if
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "SUBMIT_FLG"), 1)        
	
	
	''- �����ڰ� �����ϴ� ��쿡�� ����- �Էµ� ��� ���μ��Ű����(A100)�� �����븮�ΰ�����ȣ�� ��ġ 
	If lgcTB_54.GetData(1, "SUBMIT_FLG") = "1" Then
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "RECON_MGT_NO"), 6)
	Else
		sHTFBody = sHTFBody & UNIChar("", 6)
	End If
		
	 '- ������_����ڵ�Ϲ�ȣ :���� Check- ������ ����ڵ�Ϲ�ȣ�� ������ ����ڵ�Ϲ�ȣ ��ġ ���� Check- 
			    '�����븮�ΰ�����ȣ�� �Էµ� ��� :  ���μ��Ű����(A100)�� �����븮�� ����ڵ�Ϲ�ȣ�� ��ġ 
			    '�����븮�ΰ�����ȣ�� ������ ��� :  ���μ��Ű����(A100)�� ������ID�� ��ġ 
	if lgcTB_54.GetData(1, "RECON_MGT_NO") <>"" Then
			If  ChkNotNull(lgcTB_54.GetData(1, "AGENT_RGST_NO"), "������_����ڵ�Ϲ�ȣ") Then 
	
			      If Not  ChkNumeric(UNIRemoveDash(lgcTB_54.GetData(1, "AGENT_RGST_NO")), "������_����ڵ�Ϲ�ȣ" ) Then  blnError = True	
			   		   
			Else
			    blnError = True	
			End if    
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "AGENT_RGST_NO")), 10)
	Else
	        sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10)
	End if		
	
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "AGENT_NM"), "������_��ȣ") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "AGENT_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "REPRE_NM"), "������_����") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "REPRE_NM"), 30)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "CO_ADDR"), "������_�ּ�") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "CO_ADDR"), 80)
		
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TEL_NO"), 15)
		
	sHTFBody = sHTFBody & UNIChar("101", 3)	' ������ѱ��ڵ� 
	
	sHTFBody = sHTFBody & UNIChar("", 591)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False
	sHTFBody = ""
	
	
	' -- 1. �ں��ݺ�����Ȳ 
	sHTFBody = "B"
	sHTFBody = sHTFBody & UNIChar("40", 2)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	
	If  ChkNotNull(lgcTB_54.GetData(1, "TAX_OFFICE"), "�������ڵ�") Then 
	
	    If Not ChkNumeric(lgcTB_54.GetData(1, "TAX_OFFICE"),"�������ڵ�") Then
	      blnError = True	
	    End If  
	   
	Else
	    blnError = True	
	End If   
	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)	' -- �������ڵ� 
	
	
	

	If  ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "�Ű����_����ڵ�Ϲ�ȣ") Then 
	
	      If Not  ChkNumeric(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), "�Ű����_����ڵ�Ϲ�ȣ" ) Then  blnError = True	
			   		   
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10)
	
	
	If Not ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "�Ű����_���θ�") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "CO_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "�Ű����_��ǥ�ڼ���") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "REPRE_NM"), 30)

	' -- 2005.03.24 �������� 
'	����(���)������    
'    - ����,��00000000�����, 
'    - �ƴ� ��� ��¥���� Check, ����(���)������ >= �ι�°����⵵(��������), 
'                                ����(���)������ <= �ι�°����⵵(��������)
'    �պ�. ������    
'    - ����,��00000000�����, 
'    - �ƴ� ��� ��¥���� Check, �պ�. ������ >= �ι�°����⵵(��������), 
'                                �պ�. ������ <= �ι�°����⵵(��������)
		
	If lgcTB_54.GetData(1, "W4") <> "" Then 
		If lgcTB_54.GetData(1, "W4") <> "00000000" Then
			If Not ChkDate(lgcTB_54.GetData(1, "W4"),"����(���)������") Then
				blnError = True	
			End If

			If 	CDate(lgcTB_54.GetData(1, "W4")) = CDate(lgcTB_54.GetData(1, "W6_2"))-1 Then
			ElseIf 	CDate(lgcTB_54.GetData(1, "W4")) = CDate(lgcTB_54.GetData(1, "W6_1")) Then
			Else
				Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W4") & " , " & lgcTB_54.GetData(1, "W6_2"), "����(���)������ >= �ι�°�������(��������) �Ǵ� ����(���)������ <= �ι�°�������(��������)") 
				blnError = True
			End If
			
		End If
	    sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W4"))
	Else
	      sHTFBody = sHTFBody & UNIChar("", 8)    
	End If
	
		

	If lgcTB_54.GetData(1, "W5") <> "" Then 
		If lgcTB_54.GetData(1, "W5") <> "00000000" Then
			If Not ChkDate(lgcTB_54.GetData(1, "W5"),"�պ�.������") Then
			    blnError = True	
			End If
		End If

		If 	CDate(lgcTB_54.GetData(1, "W5")) = CDate(lgcTB_54.GetData(1, "W6_2"))-1 Then
		ElseIf 	CDate(lgcTB_54.GetData(1, "W5")) = CDate(lgcTB_54.GetData(1, "W6_1")) Then
		Else
			Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W5") & " , " & lgcTB_54.GetData(1, "W6_2"), "�պ�.������ >= �ι�°�������(��������) �Ǵ� �պ�.������ <= �ι�°�������(��������)") 
			blnError = True
		End If

	    sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W5"))
	Else
	    sHTFBody = sHTFBody & UNIChar("", 8)
	End If


' -- 2006.03.24 �������� ����	
'	If  ChkNotNull(lgcTB_54.GetData(1, "W6_1"), "�������(��������)") Then
'	    If  UNI8Date(lgcCompanyInfo.FISC_START_DT) <>UNI8Date(lgcTB_54.GetData(1, "W6_1")) Then
'	        Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W6_1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������(��������)", "���������ǻ������(��������)")) 
'			blnError = True	
'	    End If 
'	Else
'	    blnError = True	
'	End if
	
	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W6_1"))

' -- 2006.03.24 �������� ����	
'	If  ChkNotNull(lgcTB_54.GetData(1, "W6_2"), "�������(��������)") Then
'	    If  UNI8Date(lgcCompanyInfo.FISC_END_DT) <>UNI8Date(lgcTB_54.GetData(1, "W6_2")) Then
'	        Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W6_2"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�������(��������)", "���������ǻ������(��������)"))
'			blnError = True	
'	    End If 
'	Else
'	    blnError = True	
'	End if
	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W6_2"))
		
	' -- �׸��� ���� 
	
	If  ChkNotNull(lgcTB_54.GetData(2, "W10"), "�����_�ѹ����ֽļ�_������") Then 
	    '�����_�ѹ����ֽļ�_������+ �����_�ѹ����ֽļ�_�켱��	- �׸�(21)�����ֽļ�_�հ�� ��ġ 
	    dblNum1 = Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	arrTmp1(0) = lgcTB_54.GetData(2, "W11")
	arrTmp1(1) = lgcTB_54.GetData(2, "W13")
	
	Call lgcTB_54.MoveNext(2)
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "�����_�ѹ����ֽļ�_�켱��") Then blnError = True	
	  dblNum1 = dblNum1 + Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	Call lgcTB_54.Find(2, "SEQ_NO=13")	' �⸻�� �� 
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "��⸻_�ѹ����ֽļ�_������") Then blnError = True	
	dblNum2 = Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)

	arrTmp2(0) = lgcTB_54.GetData(2, "W11")
	arrTmp2(1) = lgcTB_54.GetData(2, "W13")
	
	Call lgcTB_54.MoveNext(2)
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "��⸻_�ѹ����ֽļ�_�켱��") Then blnError = True	
	dblNum2= dblNum2 + Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	If Not ChkNotNull(arrTmp1(0), "�����_�ִ�׸鰡��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp1(0), 8, 0)
	
	If Not ChkNotNull(arrTmp2(0), "��⸻_�ִ�׸鰡��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp2(0), 8, 0)
	
	If Not ChkNotNull(arrTmp1(1), "�����_�ں���") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp1(1), 16, 0)
	
	If Not ChkNotNull(arrTmp2(1), "��⸻_�ں���") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp2(1), 16, 0)
	
	Call lgcTB_54.MoveFist(2)
	Call lgcTB_54.Find(2, "SEQ_NO=3")	' ���� �������� �� 
	
	'PrintLog "SEQ_NO=3=" & lgcTB_54.EOF(2)
	'Response.End 
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W7"), "����_����1") Then blnError = True	
	If lgcTB_54.GetData(2, "W7") = "" Then 
	   sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(2, "W7"))
	Else
	   sHTFBody = sHTFBody & UNIChar("", 8)
	End If   
	      
	
	'If Not ChkBoundary("01,02,03,04,05,06,07,08,09", lgcTB_54.GetData(2, "W8"), "����_�����ڵ�1") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(2, "W8"), 2)
	if lgcTB_54.GetData(2, "W7") <> "" Then
	   '- ��¥���� Check- ���� >= ����⵵(��������), ���� <= ����⵵(��������)
	   ' ���ڰ� ����(SPACE)�� ��� �����ڵ�, ������ �����̾�� �ϰ�, �ֽļ�, �ִ�׸鰡��, �ִ����(�μ�)����, ����(����)�ں����� 0
	    sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W9"), 1, 0)   '����_����1
	Else
	    sHTFBody = sHTFBody & UNIChar("", 1)
	End If     
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "����_�ֽļ�1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W11"), "����_�ִ�׸鰡��1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W11"), 8, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W13"), "����_����(����)�ں���1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W12"), 8, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W13"), "����_����(����)�ں���1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W13"), 16, 0)
		 
	
	Call lgcTB_54.MoveNext(2)

	For iDx = 2 To 10
		If lgcTB_54.GetData(2, "W7") = "" Then 
	       sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(2, "W7"))
	    Else
	       sHTFBody = sHTFBody & UNIChar("", 8)
	   End If   
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(2, "W8"), 2)
		if lgcTB_54.GetData(2, "W7") <> "" Then
	   '- ��¥���� Check- ���� >= ����⵵(��������), ���� <= ����⵵(��������)
	   ' ���ڰ� ����(SPACE)�� ��� �����ڵ�, ������ �����̾�� �ϰ�, �ֽļ�, �ִ�׸鰡��, �ִ����(�μ�)����, ����(����)�ں����� 0
	    sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W9"), 1, 0)   '����_����1
		Else
	    sHTFBody = sHTFBody & UNIChar("", 1)
		End If  
	
		   
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W11"), 8, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W12"), 8, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W13"), 16, 0)
		
		Call lgcTB_54.MoveNext(2)
	Next
	
	' �ֽļ�������Ȳ�Ǽ� 
	iDtlCnt = lgcTB_54.RecordCount(3)
	sHTFBody = sHTFBody & UNINumeric(iDtlCnt, 6, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 6)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	
	' -- �ֽļ�������Ȳ 
	iSeqNo = 1	
	
	Do Until lgcTB_54.EOF(3) 

		sHTFBody = sHTFBody & "C"
		sHTFBody = sHTFBody & UNIChar("40", 2)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)	' -- �������ڵ� 
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10) ' -- ����ڹ�ȣ 
        If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and  UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
		     If Not ChkBoundary("1,2,3", lgcTB_54.GetData(3, "W17_1"), "���뱸��") Then blnError = True
		End if     
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W17_1"), 1)
	
	    If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
	    	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W17"), 1)
	    
	    Else
	        sHTFBody = sHTFBody & UNIChar("", 1)
	    End If	
		
	
	    If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and  UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
	    	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W18"), 40)
	    Else
	        sHTFBody = sHTFBody & UNIChar(" ", 40)
	    End If	
		if  UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then 
		    sHTFBody = sHTFBody & UNIChar("AAAAAAAAAAAAA", 13) 
		Elseif  UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 2 Then     
		    sHTFBody = sHTFBody & UNIChar("BBBBBBBBBBBBB", 13) 
		Else
		
			If  Not ChkNotNull(lgcTB_54.GetData(3, "W19"), "����_�ֹ�(�����)��Ϲ�ȣ") Then
			   blnError = True	
			    Stmp= "����_�ֹ�(�����)��Ϲ�ȣ �� " 
				Stmp = sTmp  &  " �츮���ּҰ� :	 		'CCCCCCCCCCCCC' "
				Stmp = sTmp  &  " ���ֱ����� 4, 5, 6�̰� �ش��ȣ�� �� �� ���� ����  " 
				Stmp = sTmp  &  "  1) ���δ�ü(���ֱ��У�4)�� 	'DDDDDDDDDDDDD' " 
				Stmp = sTmp  &  "    �� �ش��ȣ�� �� �� ���� ���δ�ü�� ������ �����ϴ� ���  " 
				Stmp = sTmp  &  "       'DDDDDDDDDD+�Ϸù�ȣ(3�ڸ�)'�� �����Ͽ� �Է�  " 
				Stmp = sTmp  &  "       (��)DDDDDDDDDD001, DDDDDDDDDD002,��  "
				Stmp = sTmp  &  "  2) �ܱ�������(���ֱ��У�5)�� 	'EEEEEEEEEEEEE'  "
				Stmp = sTmp  &  "    �� �ش��ȣ�� �� �� ���� �ܱ������ڰ� ������ �����ϴ� ���  " 
				Stmp = sTmp  &  "       ���� ���� ������� �����Ͽ� �Է� (��)EEEEEEEEEE001,��   "
				Stmp = sTmp  &  " 3) �ܱ�����(���ֱ��У�6)�� 	'FFFFFFFFFFFFF' " 
				Stmp = sTmp  &  "    �� �ش��ȣ�� �� �� ���� �ܱ������� ������ �����ϴ� ���  " 
				Stmp = sTmp  &  "       ���� ���� ������� �����Ͽ� �Է� (��)FFFFFFFFFF001,��   " 
				      
				Stmp = sTmp  &  "     �׿ܿ��� �����Է��ϰ�   ���� ���� �Է��� �ֽʽÿ�"
			   
			   
			   Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(3, "W19"), UNIGetMesg(Stmp, "",""))
			   
			End If    
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(3, "W19")), 13)
		End if	
	
	
	
	    '�ֽĵ����Ȳ����_�ں��ݺ�����Ȳ(84A131)�� �����_�ѹ����ֽļ��� ��ġ 
		If  ChkNotNull(lgcTB_54.GetData(3, "W20"), "�����ֽļ�") Then 
		    If dblNum1  <> UNICDbl(lgcTB_54.GetData(3, "W20"),0)  AND    UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then  '���뱸���� 1�̸� 
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, dblNum1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����ֽļ�", "�����_�ѹ����ֽļ�"))
		    End IF   
		Else
			blnError = True	
		End IF	
		
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W20"), 13, 0)
		'������ = �����ֽļ� / �ֽĵ����Ȳ����_�ں��ݺ�����Ȳ(84A131)�� �����_�ѹ����ֽļ� x 100
		If  ChkNotNull(lgcTB_54.GetData(3, "W21"), "������") Then 
		    If UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then '�հ� 
		   
					 If dblNum1 = 0 Then 
					     dblAmtRate1 = 0
					 Else
					      dblAmtRate1 = (Unicdbl(lgcTB_54.GetData(3, "W20"),0) /  dblNum1 )* 100

					 End if  
					If Unicdbl(lgcTB_54.GetData(3, "W21"),0) <> dblAmtRate1  Then 
					   blnError = True	
					   Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W21"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "������", "�����ֽļ�" & Unicdbl(lgcTB_54.GetData(3, "W20"),0) &" / �ֽĵ����Ȳ����_�ں��ݺ�����Ȳ(84A131)�� �����_�ѹ����ֽļ�" & dblNum1 &"x 100"))
                	End If 
		      End If
		   
		Else
			 blnError = True	
		End IF	 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W21"), 5, 2)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W22"), "����_���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W22"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W23"), "����_��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W23"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W24"), "����_��������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W24"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W25"), "����_���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W25"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W26"), "����_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W26"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W27"), "����_������ȯ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W27"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W28"), "����_��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W28"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W29"), "����_�絵") Then blnError = True	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W29"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W30"), "����_���") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W30"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W31"), "����_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W31"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W32"), "����_����") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W32"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W33"), "����_��Ÿ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W33"), 13, 0)
		
		'- ���뱸���� '1'(�հ�)�� ��⸻ �ѹ����ֽļ��� ��ġ�� 
		'���뱸�� '2'(�Ҿ����ּҰ�) + ���뱸�� '3'(��������)�� ��ġ �� 
		'�׸�(16)�⸻�� �׸�(11)�ֽļ��� ��ġ 
		' �⸻�ֽļ� = �����ֽļ� + ����_��� + ����_�������� + ����_�������� + ����_��� + ����_���� + ����_������ȯ + ����_��Ÿ - ����_�絵 - ����_��� - ����_���� - ����_���� - ����_��Ÿ 
		If  ChkNotNull(lgcTB_54.GetData(3, "W34"), "�⸻�ֽļ�") Then
		   ' dblNum2  -  ��⸻ �ѹ����ֽļ� 
		     If UNICDbl(lgcTB_54.GetData(3, "W34"),0) <> dblNum2  AND    UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then  '���뱸���� 1�̸� 
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W34"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�⸻�ֽļ�", "��⸻_�ѹ����ֽļ�"))
		       
		     End IF
		     
		     If Unicdbl(lgcTB_54.GetData(3, "W34"),0) <> UNICDbl(lgcTB_54.GetData(3, "W20"),0) + UNICDbl(lgcTB_54.GetData(3, "W22"),0)+ UNICDbl(lgcTB_54.GetData(3, "W23"),0) + UNICDbl(lgcTB_54.GetData(3, "W24"),0)  + UNICDbl(lgcTB_54.GetData(3, "W25"),0) _
		        + UNICDbl(lgcTB_54.GetData(3, "W26"),0) + UNICDbl(lgcTB_54.GetData(3, "W27"),0) + UNICDbl(lgcTB_54.GetData(3, "W28"),0) - UNICDbl(lgcTB_54.GetData(3, "W29"),0)  - UNICDbl(lgcTB_54.GetData(3, "W30"),0) _
		        - UNICDbl(lgcTB_54.GetData(3, "W30"),0) - UNICDbl(lgcTB_54.GetData(3, "W31"),0) - UNICDbl(lgcTB_54.GetData(3, "W32"),0) - UNICDbl(lgcTB_54.GetData(3, "W33"),0) Then
		          
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W34"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�⸻�ֽļ�", " �����ֽļ� + ����_��� + ����_�������� + ����_�������� + ����_��� + ����_���� + ����_������ȯ + ����_��Ÿ - ����_�絵 - ����_��� - ����_���� - ����_���� - ����_��Ÿ"))
		     End If
		Else     
		    blnError = True	
		End If    
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W34"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W35"), "�⸻������") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W35"), 5, 2)
		
		If lgcTB_54.GetData(3, "W36") <> "" Then
		   If Not ChkBoundary("00,01,02,03,04,05,06,07,08,09", lgcTB_54.GetData(3, "W36"), "�������ֿ��ǰ���") Then blnError = True
		   If (lgcTB_54.GetData(3, "W17") = "2" Or lgcTB_54.GetData(3, "W17") = "3" Or lgcTB_54.GetData(3, "W17") = "4" Or lgcTB_54.GetData(3, "W17") = "6" ) And lgcTB_54.GetData(3, "W36") <> "09" Then 
		      '���ֱ����� ��2��,��3��,��4��,��6�� �̸� �������ֿ��� ����� ��09���� �ƴϸ� ���� 
		        blnError = True	
			    Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W36"),0), UNIGetMesg("���ֱ����� ��2��,��3��,��4��,��6�� �̸� �������ֿ��� ����� ��09���� �ƴϸ� �����Դϴ�", "", ""))
		   End IF
		End If   
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W36"), 2)
	
			
		sHTFBody = sHTFBody & UNIChar("", 519) & vbCrLf

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_54.MoveNext(3)	' -- 1�� ���ڵ�� 
	Loop



	PrintLog "Write2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_54 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9109MA1 : " & lgStrSQL
End Sub
%>
