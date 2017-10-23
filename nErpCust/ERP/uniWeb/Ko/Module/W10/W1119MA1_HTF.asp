<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��3ȣ �����׿���ó��(��ձ�ó��)
'*  3. Program ID           : W1119MA1
'*  4. Program Name         : W1119MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_3_3_4

Set lgcTB_3_3_4 = Nothing ' -- �ʱ�ȭ 

Class C_TB_3_3_4
	' -- ���̺��� �÷����� 
	Dim W_TYPE
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W8
	Dim W10
	Dim W11
	Dim W12
	Dim W13
	Dim W14
	Dim W15
	Dim W16
	Dim W17
	Dim W18
	Dim W19
	Dim W20
	Dim W25
	
	' -- ���� 2006.03
	Dim W26	
	Dim W27
	Dim W28
	' ---
	
	Dim W30
	Dim W31
	Dim W32
	Dim W33
	Dim W34
	Dim W35
	Dim W40
	Dim W41
	Dim W42
	Dim W43
	Dim W44
	Dim W50
	Dim W_DT
	
	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
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

		' -- ��1ȣ������ �о�´�.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '�� : Make sql statements

		' -- Ŀ���� Ŭ���̾�Ʈ�� ���� **���� ../wcm/incServerADoDb.asp ���� �����Ǵ� ��� 
		gCursorLocation = adUseClient 
		
			If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
				If Not CALLED_OUT Then	' -- �ܺο��� �θ� ���� ȣ�����ʿ��� ����Ÿ������ �����Ѵ�. ������ lgsPGM_ID, lgsPGM_NM�� ȣ���ѳ��̱⶧���̴�.
				    IF CHK_COMPANY = TRUE THEN
					   Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
					 END IF	  
				End If
			    Exit Function
			End If
		

		' ��Ƽ�������� ù���� ���� 
		Call GetData
		
		Call CloseRs
		
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
	
	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W_TYPE		= lgoRs1("W_TYPE")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			W8			= lgoRs1("W8")
			W10			= lgoRs1("W10")
			W11			= lgoRs1("W11")
			W12			= lgoRs1("W12")
			W13			= lgoRs1("W13")
			W14			= lgoRs1("W14")
			W15			= lgoRs1("W15")
			W16			= lgoRs1("W16")
			W17			= lgoRs1("W17")
			W18			= lgoRs1("W18")
			W19			= lgoRs1("W19")
			W20			= lgoRs1("W20")
			W25			= lgoRs1("W25")
			
			' -- ���� 2006.03
			W26			= lgoRs1("W26")
			W27			= lgoRs1("W27")
			W28			= lgoRs1("W28")
			
			
			W30			= lgoRs1("W30")
			W31			= lgoRs1("W31")
			W32			= lgoRs1("W32")
			W33			= lgoRs1("W33")
			W34			= lgoRs1("W34")
			W35			= lgoRs1("W35")
			W40			= lgoRs1("W40")
			W41			= lgoRs1("W41")
			W42			= lgoRs1("W42")
			W43			= lgoRs1("W43")
			W44			= lgoRs1("W44")
			W50			= lgoRs1("W50")
			W_DT		= lgoRs1("W_DT")
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
				lgStrSQL = lgStrSQL & " FROM TB_3_3_4 A WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W1119MA1
	Dim A100_BASIC
	Dim A100
	Dim A113
	Dim A115
	Dim A110
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W1119MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, strMAIN_IND, dblCode5678 ,dblCode8273
    
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1119MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1119MA1"

	Set lgcTB_3_3_4 = New C_TB_3_3_4		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_3_3_4.LoadData Then Exit Function			
    Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
	
	'==========================================
	' -- ��3ȣ �����׿���ó��(��ձ�ó��) �������� 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 

	If Not ChkNotNull(lgcTB_3_3_4.W_DT, "ó��Ȯ����") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_3_3_4.W_DT)


	If lgcCompanyInfo.Comp_type2 = "2" then   '���������ΰ�� 
	   '�ڵ�(10)�հ谡 �����̰�, �ڵ�(15)������ ��0���� �ƴϸ�I.ó������ձ� �׸���� ���� I.ó���������׿��� �׸�鿡 �Է��ؾ� ��.
	  ' �� �ܿ��� �ڵ�(10)�հ谡 ������ ���I.ó���������׿��� �׸���� ���� I.ó������ձ� �׸�鿡 �Է��ؾ� ��.
    
		 If unicdbl(lgcTB_3_3_4.W10,0) < 0  and  unicdbl(lgcTB_3_3_4.W15,0) <> 0 Then
			Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg("���������ΰ�� �հ�(10)�� �����̰� ������(15)�� 0�̾ƴϸ� 1.ó���� ��ձ� �׸񰪵��� ���� ó���������׿��� �׸�鿡 �Է��ؾ���", "",""))
			blnError = True	
		  END IF
	Else	  	
	     if unicdbl(lgcTB_3_3_4.W10,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg("ó���������׿��� �׸���� ���� ó���� ��ձ� �׸�鿡 �Է��ؾ���", "",""))
			blnError = True	
		  END IF
	End If



	'* �ڵ� 02 + 03 + 04 - 05 + 06
	'*ǥ�ش�������ǥ�� ó���������׿��� �Ǵ� ó������ձݰ� ��ġ  
	 ' �Ϲݹ��� : ǥ�ش�������ǥ(�Ϲݹ���)(A113) �� �ڵ� (56)               ó���������׿��ݶǴ�ó������ձ� 
	 ' �������� : ǥ�ش�������ǥ(��������)(A114) �� �ڵ� (78)              ó���������׿��ݶǴ�ó������ձ� 


	If  ChkNotNull(lgcTB_3_3_4.W1, "ó���������׿���") Then 

		if lgcCompanyInfo.Comp_type2 = "1" then 
		    dblCode5678 =  Getdata_TB_3_3_4_A142("A113_1")
		   
		Else
		    dblCode5678=  Getdata_TB_3_3_4_A142("A114_1")
		End if   

		if unicdbl(dblCode5678,0) <> unicdbl(lgcTB_3_3_4.W1,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ó���������׿���","ǥ�ش�������ǥ�� ó���������׿��� �Ǵ� ó������ձ�"))
		    blnError = True	
		End if
		
		if unicdbl( lgcTB_3_3_4.W1,0)  <> unicdbl( lgcTB_3_3_4.W2,0) + unicdbl( lgcTB_3_3_4.W3,0) +  unicdbl( lgcTB_3_3_4.W4,0) -  unicdbl( lgcTB_3_3_4.W5,0) +  unicdbl( lgcTB_3_3_4.W6,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ó���������׿���","�ڵ� 02 + 03 + 04 - 05 + 06"))
		    blnError = True	
		End if
		
		
		
		
			
			
	Else
	        blnError = True	
	End if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W1, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W2, "�����̿������׿���(�Ǵ� �����̿���ձ�)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W2, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W3, "ȸ�躯���� ����ȿ��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W3, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W4, "���������������(�Ǵ� ������������ս�)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W4, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W5, "�߰�����") Then 
	    if unicdbl(lgcTB_3_3_4.W5,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W5, UNIGetMesg("�߰����� ���� �����Դϴ�.", "",""))
	         blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W5, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W6, "��������(�Ǵ� �����ս�)") Then 
	  ' - �Ϲݹ��� : ǥ�ؼ��Ͱ�꼭(�Ϲݹ���)(A115)�� �ڵ�(82)��������(���ս�)
	  '- �������� : ǥ�ؼ��Ͱ�꼭(��������)(A116)�� �ڵ�(73)��������(���ս�)
	   if lgcCompanyInfo.Comp_type2 = "1" then 
		    dblCode8273 =  Getdata_TB_3_3_A142("A115_1")
		Else
		    dblCode8273=  Getdata_TB_3_3_A142("A116_1")
		End if 
		
		if unicdbl(dblCode8273,0) <> unicdbl(lgcTB_3_3_4.W6,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��������","ǥ�ش�������ǥ�� ��������(���ս�)"))
		    blnError = True	
		End if
	
	Else
		blnError = True	
    End if		
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W6, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W8, "���������� ���� ���Ծ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W8, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W10, "�հ�") Then
	    '- �ڵ�(01)ó���������׿��� + �ڵ�(08)���������� ���� ���Ծ� 
	     if unicdbl( lgcTB_3_3_4.W10,0)  <> unicdbl( lgcTB_3_3_4.W1,0) + unicdbl( lgcTB_3_3_4.W8,0)  then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�հ�","ó���������׿���(01) + ���������� ���� ���Ծ�(08)"))
		    blnError = True	
		 End if
	Else
	     blnError = True	
	End if	 
		
		
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W10, 15, 0)


	If  ChkNotNull(lgcTB_3_3_4.W11, "�����׿��� ó�о�") Then
	     '- �ڵ� 12 + 13 + 14 + 15 + 18 + 19 + 20 + 26 + 27 + 28 (2006.03����)
	     if unicdbl( lgcTB_3_3_4.W11,0)  <> unicdbl( lgcTB_3_3_4.W12,0) + unicdbl( lgcTB_3_3_4.W13,0) + unicdbl( lgcTB_3_3_4.W14,0) + unicdbl( lgcTB_3_3_4.W15,0) + unicdbl( lgcTB_3_3_4.W18,0) + unicdbl( lgcTB_3_3_4.W19,0) + unicdbl( lgcTB_3_3_4.W20,0) + unicdbl( lgcTB_3_3_4.W26,0) + unicdbl( lgcTB_3_3_4.W27,0) + unicdbl( lgcTB_3_3_4.W28,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W11, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����׿��� ó�о�","ó���������׿���(01) + ���������� ���� ���Ծ�(08)"))
		    blnError = True	
		 End if
	Else
	     blnError = True	
	End if	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W11, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W12, "�����غ��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W12, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W13, "��Ÿ����������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W13, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W14, "�ֽ����ι������ݻ󰢾�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W14, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W15, "����") Then 
	    '�ڵ� 16 + 17
	    if unicdbl( lgcTB_3_3_4.W15,0)  <> unicdbl( lgcTB_3_3_4.w16,0) + unicdbl( lgcTB_3_3_4.w17,0)  then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W15, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "����","��.���ݹ��(16) + ��.�ֽĹ��(17)"))
		    blnError = True	
		 End if
	Else
	    blnError = True	
	End if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W15, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W16, "���ݹ��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W16, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W17, "�ֽĹ��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W17, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W18, "���Ȯ��������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W18, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W19, "��ä������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W19, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W20, "��Ÿ������") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W20, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W25, "�����̿������׿���") Then 
	     '�ڵ� 01 + 08 + 11
	    if unicdbl( lgcTB_3_3_4.W25,0)  <> unicdbl( lgcTB_3_3_4.w1,0) + unicdbl( lgcTB_3_3_4.w8,0) -unicdbl( lgcTB_3_3_4.w11,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W25, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����̿������׿���","ó���������׿���(01) + ���������� ���� ���Ծ�(08) - �����׿��� ó�о�(11)"))
		    blnError = True	
		 End if
	 
	Else
	    blnError = True	
	End if  
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W25, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W30, "ó������ձ�") Then
	    If unicdbl(lgcTB_3_3_4.W30,0) <> 0 Then
				If unicdbl(lgcTB_3_3_4.W30,0) < 0 then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg("ó������ձ� ���� �����Դϴ�.", "",""))
				     blnError = True	
				End if
	    
	    
				'�ڵ� (31 + 32 + 33 - 34 - 35) X (-1)
				If(unicdbl( lgcTB_3_3_4.w31,0) + unicdbl( lgcTB_3_3_4.w32,0) + unicdbl( lgcTB_3_3_4.w33,0)-unicdbl( lgcTB_3_3_4.w34,0)-unicdbl( lgcTB_3_3_4.w35,0)) * -1 < 0 and unicdbl(lgcTB_3_3_4.W30,0) <> 0 then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg("", "�ڵ� (31 + 32 + 33 - 34 - 35) X (-1)�� ����� �����̹Ƿ� �����׿���ó�а�꼭�� �Է��ϴ� ����̰ų� �ڵ�(31,35)���� ǥ���� �����Դϴ�"))
				    blnError = True	
				 End If
				 
				 
				if lgcCompanyInfo.Comp_type2 = "1" then 
				    dblCode5678 =  Getdata_TB_3_3_4_A142("A113_1")
				   
				Else
				    dblCode5678=  Getdata_TB_3_3_4_A142("A114_1")
				End if   

				if abs(unicdbl(dblCode5678,0)) <> unicdbl(lgcTB_3_3_4.W30,0) then
				   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "ó������ձ�","ǥ�ش�������ǥ�� ó���������׿��� �Ǵ� ó������ձ�"))
				    blnError = True	
				End if
		 End If		
	
	    
	Else
		 blnError = True	
	End If	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W30, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W31, "�����̿������׿���(�Ǵ� �����̿���ձ�)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W31, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W32, "ȸ�躯���� ����ȿ��") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W32, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W33, "���������������(�Ǵ� ������������ս�)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W33, 15, 0)


	
	If  ChkNotNull(lgcTB_3_3_4.W34, "�߰�����") Then 
	    if unicdbl(lgcTB_3_3_4.W34,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W34, UNIGetMesg("�߰����� ���� �����Դϴ�.", "",""))
	         blnError = True	
	    End if
	Else
	    blnError = True	
	End if   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W34, 15, 0)
	

	If  ChkNotNull(lgcTB_3_3_4.W35, "�����ս�(�Ǵ� ��������)") Then 
	    ' - �Ϲݹ��� : ǥ�ؼ��Ͱ�꼭(�Ϲݹ���)(A115)�� �ڵ�(82)��������(���ս�)
	  '- �������� : ǥ�ؼ��Ͱ�꼭(��������)(A116)�� �ڵ�(73)��������(���ս�)
	  IF unicdbl(lgcTB_3_3_4.W35,0) <> 0 Then
			if lgcCompanyInfo.Comp_type2 = "1" then 
				    dblCode8273 =  Getdata_TB_3_3_A142("A115_1")
				Else
				    dblCode8273=  Getdata_TB_3_3_A142("A116_1")
				End if 
		
				if unicdbl(dblCode8273,0)*(-1) <> unicdbl(lgcTB_3_3_4.W35,0)  then
				   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W35, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����ս�(�Ǵ� ��������","ǥ�ش�������ǥ�� ��������(���ս�)"))
				    blnError = True	
				End if
		End If		
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W35, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W40, "��ձ�ó����") Then 
	
	    If unicdbl(lgcTB_3_3_4.W40,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg("��ձ�ó���� ���� �����Դϴ�.", "",""))
	         blnError = True	
	    End if
	    
	    
	     '-41 + 42 + 43 + 44
	     if unicdbl( lgcTB_3_3_4.W40,0)  <> unicdbl( lgcTB_3_3_4.W41,0) + unicdbl( lgcTB_3_3_4.W42,0) + unicdbl( lgcTB_3_3_4.W43,0) + unicdbl( lgcTB_3_3_4.W44,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "��ձ�ó����","�ڵ�(41 + 42 + 43 + 44)"))
		    blnError = True	
		 End if
	Else
		blnError = True	
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W40, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W41, "�������������Ծ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W41, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W42, "��Ÿ�������������Ծ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W42, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W43, "�����غ�����Ծ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W43, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W44, "�ں��׿������Ծ�") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W44, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W50, "�����̿���ձ�") Then 
	   '30-40
	     if unicdbl( lgcTB_3_3_4.W50,0)  <>  unicdbl( lgcTB_3_3_4.W30,0) - unicdbl( lgcTB_3_3_4.W40,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W50, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�����̿���ձ�","�ڵ�(30-40)"))
		    blnError = True	
		 End if
	Else
	    blnError = True	
	End If 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W50, 15, 0)
	
	
	
	if unicdbl(lgcTB_3_3_4.W25,0) <> 0 and unicdbl( lgcTB_3_3_4.W50,0) <> 0 then
	   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg("�����̿������׿���(25)�� �����̿� ��ձ�(50) ��� �ݾ��� �Է��� �� �����ϴ�","", ""))
		    blnError = True	
	End if

	' -- 2006.03 ���� 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W26, 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W27, 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W28, 15, 0)

	' -- 47ȣ �� ������ �ڵ�98 �ݾװ� ��ġ �� �ʿ��� : 2006.03 
	If lgcTB_3_3_4.W_TYPE = "1" Then
		If UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0) > 0 And lgcCompanyInfo.Comp_type2 <> "2" Then	' 2006.03.24 �������� ���α����� 2�� �ƴҶ� 
			dblCode5678 =  Getdata_TB_47_A110("A110")

			if dblCode5678 <> UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0)  then
			   Call SaveHTFError(lgsPGM_ID, UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0)  & " <> " &dblCode5678 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(05)�߰����� + �ڵ�(15)���� + �ڵ�(26)����ó�п����ѻ󿩱�", "��47ȣ �ֿ��������(��)(A110)�� �ڵ�(98) ����ó�бݾ�"))
			   blnError = True	
			End if

		End If
	Else
		If UNICdbl(lgcTB_3_3_4.W34, 0) > 0  And lgcCompanyInfo.Comp_type2 <> "2" Then	' 2006.03.24 �������� ���α����� 2�� �ƴҶ� 
			dblCode5678 =  Getdata_TB_47_A110("A110")

			if dblCode5678 <> UNICdbl(lgcTB_3_3_4.W34, 0)  then
			   Call SaveHTFError(lgsPGM_ID, UNICdbl(lgcTB_3_3_4.W34, 0) & " <> " &dblCode5678 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(34)�߰�����", "��47ȣ �ֿ��������(��)(A110)�� �ڵ�(98) ����ó�бݾ�"))
			   blnError = True	
			End if
		End If
	
	End If	


	sHTFBody = sHTFBody & UNIChar("", 26)	' -- ���� 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3_3_4 = Nothing	' -- �޸����� 
	
End Function





Function CHK_COMPANY()
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , chkData1,chkData2,chkData3, chkData4
 CHK_COMPANY = FALSE

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A100   = new C_TB_1	' -- W8101MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
        cDataExists.A100.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A100.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ��			
						
      
						
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��1ȣ ����", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
          
		    chkData1 =cDataExists.A100.W2                 '�������������� 
		
		End If	
						
		
		Set cDataExists.A100 = Nothing
		
		



		chkData2 =lgcCompanyInfo.COMP_TYPE1							'���α��� 
		chkData3 =lgcCompanyInfo.HOME_TAX_MAIN_IND					'�������ȣ 
		chkData4 =mid(replace(lgcCompanyInfo.OWN_RGST_NO,"-",""),4,2)'����ڹ�ȣ(4,2)
					
		

		
		Set cDataExists = Nothing	' -- �޸����� 
		

		
	   If (chkData1 = "60" Or chkData1 = "70" ) and chkData3 = "999999" then
	      CHK_COMPANY = False
	      Exit Function
	   End if   
	   
	   
	     If chkData2 = "2" and  chkData4 = "84" then
	      CHK_COMPANY = False
	      Exit Function
	   End if 

 CHK_COMPANY = TRUE
End Function

Function Getdata_TB_3_3_4_A142(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A113   = new C_TB_3_2	' -- W1101MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
        cDataExists.A113.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A113.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ��			
		
		
	
		   Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   '�Ϲݹ��� 
		  	
		   cDataExists.A113.WHERE_SQL  =  lgStrSQL			
		
		If Not cDataExists.A113.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " ǥ�ش�������ǥ", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
          
		    dblData =unicdbl(cDataExists.A113.CR_INV,0)          
		
		End If	
						
		
		Set cDataExists.A113 = Nothing
		
		
		
		Set cDataExists = Nothing	' -- �޸����� 
		
		 Getdata_TB_3_3_4_A142 = dblData
	
End Function



Function Getdata_TB_3_3_A142(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A115   = new C_TB_3_3	' -- W1101MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
        cDataExists.A115.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A115.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ��			
		
		
		   Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   '�Ϲݹ��� 
		  				

		   cDataExists.A115.WHERE_SQL  =  lgStrSQL				
		If Not cDataExists.A115.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " ǥ�ؼ��Ͱ�꼭", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
          
		    dblData =unicdbl(cDataExists.A115.W5,0)          
		
		End If	
						
		
		Set cDataExists.A115 = Nothing
		
		
		
		Set cDataExists = Nothing	' -- �޸����� 
		
		 Getdata_TB_3_3_A142 = dblData
	
End Function

' -- 2006.03 �����߰� 
Function Getdata_TB_47_A110(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A110   = new C_TB_47A	' -- W9101MA1_HTF.asp �� ���ǵ� 
							
		' -- �߰� ��ȸ������ �о�´�.
        cDataExists.A110.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
		cDataExists.A110.WHERE_SQL = ""			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ��			
		
		Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   
		  	
		   cDataExists.A110.WHERE_SQL  =  lgStrSQL			
		
		If Not cDataExists.A110.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " ��47ȣ �ֿ��������(��) ", TYPE_DATA_NOT_FOUND)		' -- �ܺο��� ȣ�������� ����Ÿ������ ��������� �Ѵ�.
		Else
          
		    dblData =unicdbl(cDataExists.A110.W125,0)          
		
		End If	
						
		
		Set cDataExists.A110 = Nothing
		
		Set cDataExists = Nothing	' -- �޸����� 
		
		 Getdata_TB_47_A110 = dblData
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W1119MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A113_1" '-- �ܺ� ���� SQL
	       lgStrSQL =""
	       lgStrSQL = " and  par_gp_cd = '33'  and A.W4 = '56'  "
	  
	  
	  Case "A114_1" '-- �ܺ� ���� SQL
	       lgStrSQL =""
	       lgStrSQL = " and par_gp_cd = '33'  and  A.W4 = '78' "    
	       
	  Case "A115_1" '-- �ܺ� ���� SQL
	       lgStrSQL =""
	       lgStrSQL = " and A.W4 = '82'  "
	  
	  
	  Case "A116_1" '-- �ܺ� ���� SQL
	       lgStrSQL =""
	       lgStrSQL = " and  A.W4 = '73' "            
   
	  Case "A110"
		   lgStrSQL =""
	End Select
	PrintLog "SubMakeSQLStatements_W1119MA1 : " & lgStrSQL
End Sub
%>
