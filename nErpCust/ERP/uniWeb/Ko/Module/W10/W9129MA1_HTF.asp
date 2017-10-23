
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : �ؿ��������� ���� 
'*  3. Program ID           : W9129MA1
'*  4. Program Name         : W9129MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : lee wol san
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_A128

Set lgcTB_A128 = Nothing ' -- �ʱ�ȭ 

Class C_TB_A128
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
				lgStrSQL = lgStrSQL & " FROM TB_A128	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W6 <> '' "  & vbCrLf	' -- ����Ÿ�� ���� ���� 
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9129MA1
	Dim A126
End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9129MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), blnChkA126A127
    DIM oRs3,oRs4
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    
    PrintLog "MakeHTF_W9129MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9129MA1"

	Set lgcTB_A128 = New C_TB_A128		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_A128.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9129MA1

	' -- �������� 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' �۷ι����� ���۴��ڵ� 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' ������� 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' �Ű��� 

	'==========================================
	' -- ��4ȣ �����Ѽ�������꼭 �������� 
	iSeqNo = 1	: sHTFBody = ""

	Do Until lgcTB_A128.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 


		'3 ��������ȣ 
		If Not ChkNotNull(lgcTB_A128.GetData("SEQ_NO"), "��������ȣ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("SEQ_NO"), 4, 0)
		
		'4 ���⸻��������� 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W15"), 5, 0)
		
		'5 ���ż������ 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W16"), 5, 0)
		
		'6 ����������� 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W17"), 5, 0)
		
		'7 ��⸻��������� 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W18"), 5, 0)
		
		'8 �������� 
		If Not ChkNotNull(lgcTB_A128.GetData("W6"), "��ī�ڵ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W6"), 3)
		
		'9 �ؿ������ 
		If Not ChkNotNull(lgcTB_A128.GetData("W7"), "���������") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W7"), 60)
		
		'10 �ؿ����������ȣ 
		If Not ChkNotNull(lgcTB_A128.GetData("W8"), "�������������ȣ") Then blnError = True	
		If Len(lgcTB_A128.GetData("W8")) <> 8 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W8"), UNIGetMesg("��ü���̰� 8�� �ƴϸ� �����Դϴ�.", "",""))
		End If
		
		If Left(lgcTB_A128.GetData("W8"), 1) <> "9" And Left(lgcTB_A128.GetData("W8"), 1) <> "8" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W8"), UNIGetMesg("ù���ڰ� 9, 8 ��(��) �ƴϸ� �����Դϴ�", "",""))
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W8"), 8)
		
		
		'11 �Ǹ����� 
		If Not ChkBoundary("1,2", lgcTB_A128.GetData("W9"), "��������: " & lgcTB_A128.GetData("W9") & " " ) Then blnError = True
      
		Call SubMakeSQLStatements_W9129MA1("A",lgcTB_A128.GetData("SEQ_NO"),iKey1,iKey2, iKey3)  '
        If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("�ؿ����� �濵��Ȳ�� �Է��ϼ���.", "",""))
		End If
		if lgcTB_A128.GetData("W9")="1" then '���� 

			if oRs3("allSum")="0" then '��� sum�� 0 �̸� ���� 
				Call SaveHTFError(lgsPGM_ID, oRs3("allSum"), UNIGetMesg("������ ��� �ڻ��Ѱ�~���������� �ݾ��� 0�̸� �����Դϴ�. ", "",""))
			end if
		elseif  lgcTB_A128.GetData("W9")="2" then '�繫�� 
			if oRs3("allSum")<>"0" then '��� sum�� 0 <> �̸� ���� 
				Call SaveHTFError(lgsPGM_ID, oRs3("allSum"), UNIGetMesg("�繫���� ��� �ڻ��Ѱ�~���������� �ݾ��� 0�̾�� �մϴ�. ", "",""))
			end if
		else
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W9"), UNIGetMesg("1 �Ǵ� 2 �̾�� �մϴ�. ", "",""))
		end if
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W9"), 1)
		
		'12 �������� 
		If Not ChkNotNull(lgcTB_A128.GetData("W10"), "��������") Then blnError = True
		If DateDiff("m", lgcTB_A128.GetData("W10"), lgcTB_A128.GetData("W21")) < 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W10"), UNIGetMesg("�������ڴ� ����Ϻ��� �۾ƾ��մϴ�.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A128.GetData("W10"))
		
		'13 �ؿ���������� 
		If Not ChkNotNull(lgcTB_A128.GetData("W11"), "�������� ������") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W11"), 70)
		
		'14 �����ڵ� 
		 Call SubMakeSQLStatements_W9129MA1("2",lgcTB_A128.GetData("W12"), "", "")  '�����ڵ� 
		
		If   FncOpenRs("R",lgObjConn, oRs4, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("�����ڵ尡 �������� �ʽ��ϴ�.", "",""))
		End If
			Call SubCloseRs(oRs4)
		If Not ChkNotNull(lgcTB_A128.GetData("W12"), "�����ڵ�") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W12"), 7)


		'15 ������ 
		If Not ChkNotNull(lgcTB_A128.GetData("W13"), "������") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W13")), "������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W13"), 5, 0)
		
		'16 �����İ������� 
		
		If Not ChkNotNull(lgcTB_A128.GetData("W14"), "�����İ�������") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W14")), "�����İ�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W14"), 5, 0)
		
		


		
		'=========================================================================
		
		'17 �ڻ��Ѱ� 
		If Not ChkNotNull(oRs3("W1"), "�ڻ��Ѱ�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W1")), "�ڻ��Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W1"), 15, 0)

		'18 �����װ��๰ 
		If Not ChkNotNull(oRs3("W2"), "�����װ��๰") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W2")), "�����װ��๰") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W2"), 15, 0)

		'19 �����ġ,������ݱ� 
		If Not ChkNotNull(oRs3("W3"), "�����ġ,������ݱ�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W3")), "�����ġ,������ݱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W3"), 15, 0)
	
		'20 �ڻ��Ÿ 
		If Not ChkNotNull(oRs3("W4"), "�ڻ��Ÿ") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W4")), "�ڻ��Ÿ") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W4"), 15, 0)

		'21 ��ä�Ѱ� 
		If Not ChkNotNull(oRs3("W5"), "��ä�Ѱ�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W5")), "��ä�Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W5"), 15, 0)

		'22 �ں��Ѱ� 
		If Not ChkNotNull(oRs3("W6"), "�ں��Ѱ�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W6")), "�ں��Ѱ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W6"), 15, 0)


		'23 ����������� 
		If Not ChkNotNull(oRs3("W7"), "�����������") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W7")), "�����������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W7"), 15, 0)


		'24 ����� 
		If Not ChkNotNull(oRs3("W8"), "�����") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W8")), "�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W8"), 15, 0)

		'25 ������� 
		If Not ChkNotNull(oRs3("W9"), "�������") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W9")), "�������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W9"), 15, 0)

		'26 �Ǹź���Ϲݰ����� 
		If Not ChkNotNull(oRs3("W10"), "�Ǹź���Ϲݰ�����") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W10")), "�Ǹź���Ϲݰ�����") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W10"), 15, 0)

		'27 �����ܼ��� 
		If Not ChkNotNull(oRs3("W11"), "�����ܼ���") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W11")), "�����ܼ���") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W11"), 15, 0)

		'28 �����ܺ�� 
		If Not ChkNotNull(oRs3("W12"), "�����ܺ��") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W12")), "�����ܺ��") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W12"), 15, 0)

		'29 Ư������ 
		If Not ChkNotNull(oRs3("W13"), "Ư������") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W13")), "Ư������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W13"), 15, 0)

		'30 Ư���ս� 
		If Not ChkNotNull(oRs3("W14"), "Ư���ս�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W14")), "Ư���ս�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W14"), 15, 0)

		'31 ���μ� 
		If Not ChkNotNull(oRs3("W15"), "���μ�") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W15")), "���μ�") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W15"), 15, 0)

		'32 �������� 
		If Not ChkNotNull(oRs3("W16"), "��������") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W16")), "��������") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W16"), 15, 0)
		Call SubCloseRs(oRs3)
		'=========================================================================
		'33 ����� 
		If CDbl(lgcTB_A128.GetData("W22")) > 0 Then
			If Not ChkNotNull(lgcTB_A128.GetData("W21"), "�����") Then 
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W21"), UNIGetMesg("ȸ���ݾ��� 0���� ũ�� ������� �ݵ�� �ԷµǾ�� �մϴ�.", "",""))
			End If
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A128.GetData("W21"))

		'34 ȸ���ݾ׿�ȭ 
		
		If Not ChkNotNull(lgcTB_A128.GetData("W22"), "ȸ���ݾ׿�ȭ") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W22")), "ȸ���ݾ׿�ȭ") Then blnError = True		
		If IsDate(lgcTB_A128.GetData("W21")) Then
			If CDbl(lgcTB_A128.GetData("W22")) <= 0 Then
				blnError = True	
				Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W22"), UNIGetMesg("������� �ԷµǸ� ȸ���ݾ��� 0���� Ŀ�� �մϴ�.", "",""))
			End If
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W22"), 15, 0)
		'35 ���� 
		sHTFBody = sHTFBody & UNIChar("", 40) ' -- ����	 :
		If Not blnError Then
			Call WriteLine2File(sHTFBody)
		End If
		sHTFBody=""
		
		lgcTB_A128.MoveNext 
	Loop

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.

	
	
	If Not blnError Then
		'Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_A128 = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9129MA1(pMode,pVal, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A126" '-- �ܺ� ���� SQL

			lgStrSQL = ""
			
      Case "A"
      
		lgStrSQL =  " select max(w1) w1,max(w2 ) w2 ,max(w3 ) w3 ,max(w4 ) w4 ,max(w5 ) w5 ,max(w6 ) w6 ," & CHR(13)
		lgStrSQL = lgStrSQL & " max(w7 ) w7 ,max(w8 ) w8 ,max(w9 ) w9 ,max(w10) w10,max(w11) w11,max(w12) w12,    " & CHR(13)
		lgStrSQL = lgStrSQL & " max(w13) w13,max(w14) w14,max(w15) w15,max(w16) w16, sum(allSum) allSum                               " & CHR(13)
		lgStrSQL = lgStrSQL & " from (                                                                            " & CHR(13)
		lgStrSQL = lgStrSQL & " select  case when w2='01' then max(w3) end w1,                                    " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='02' then max(w3) end w2 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='03' then max(w3) end w3 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='04' then max(w3) end w4 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='05' then max(w3) end w5 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='06' then max(w3) end w6 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='07' then max(w3) end w7 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='08' then max(w3) end w8 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='09' then max(w3) end w9 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='10' then max(w3) end w10,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='11' then max(w3) end w11,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='12' then max(w3) end w12,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='13' then max(w3) end w13,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='14' then max(w3) end w14,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='15' then max(w3) end w15,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='16' then max(w3) end w16,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  sum(w3) allSum																	  " & CHR(13)

		lgStrSQL = lgStrSQL & " from  tb_a128_1                                                                   " & CHR(13)
		lgStrSQL = lgStrSQL & " where  tb_a128_1.co_cd=" &  pCode1   & "                                          " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.fisc_year=" &  pCode2   & "                                         " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.rep_type=" &  pCode3   & "                                                    " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.seq_no = '"&pVal&"'                                                        " & CHR(13)
		lgStrSQL = lgStrSQL & " group by w2 ) a                                                                   " & CHR(13)

	Case "2" '����üũ 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " select top 1 STD_INCM_RT_CD from tb_std_income_rate"  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE ATTRIBUTE_YEAR = '2005' " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND STD_INCM_RT_CD= " & filterVar(pCode1 ,"''","S")	 & vbCrLf
			
	End Select
	PrintLog "SubMakeSQLStatements_W9129MA1 : " & lgStrSQL
End Sub
%>
