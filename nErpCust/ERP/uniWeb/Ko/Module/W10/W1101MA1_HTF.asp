
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  ��3ȣ��2 ǥ�ش�������ǥ 
'*  3. Program ID           : W1101MA1
'*  4. Program Name         : W1101MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_3_2

Set lgcTB_3_2 = Nothing ' -- �ʱ�ȭ 

Class C_TB_3_2
	' -- ���̺��� �÷����� 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim DR_INV
	Dim CR_INV
	
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
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		' ��Ƽ�������� ù���� ���� 
		Call GetData
		
		LoadData = True
	End Function

	'----------- ��Ƽ �� ���� ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
        Call GetData
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function Clone(Byref pRs)
		Set pRs = lgoRs1.clone
	End Function
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			DR_INV		= lgoRs1("DR_INV")
			CR_INV		= lgoRs1("CR_INV")
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
				lgStrSQL = lgStrSQL & "SELECT A.W1, A.W2, ( CASE WHEN  LEFT(A.W2, 1) = '1' THEN B.W5 ELSE 0 END ) AS DR_INV, A.W5 " & vbCrLf
				lgStrSQL = lgStrSQL & ",  A.W3, A.W4, A.W6 " & vbCrLf
				lgStrSQL = lgStrSQL & ", (CASE WHEN LEFT(A.W2, 1) <> '1' THEN B.W5 ELSE 0 END ) AS CR_INV " & vbCrLf
				lgStrSQL = lgStrSQL & "FROM TB_3_2_2 A (NOLOCK)   " & vbCrLf
				lgStrSQL = lgStrSQL & "	INNER JOIN TB_3_2_1 B (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.W1=B.W1 AND A.W2=B.W2 " & vbCrLf 
				lgStrSQL = lgStrSQL & "	INNER JOIN TB_COMPANY_HISTORY C (NOLOCK) ON A.CO_CD=C.CO_CD AND A.FISC_YEAR=C.FISC_YEAR AND A.REP_TYPE=C.REP_TYPE AND A.W1=C.COMP_TYPE2  " & vbCrLf 
				lgStrSQL = lgStrSQL & "	LEFT OUTER JOIN dbo.ufn_TB_ACCT_GP('200503') D ON A.W1=D.COMP_TYPE2 AND A.W2=D.GP_CD  " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 
				
				lgStrSQL = lgStrSQL & "  ORDER BY  LEFT(A.W2, 1), D.gp_seq " & vbCrLf
  				
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W1101MA1
	Dim A144

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W1101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, blnFirst, blnFile,dblAmt1 ,dblAmt72,dblAmt66,dblAmt43,dblAmt84
    
    Const TYPE_1 = 0	' ǥ�ش�������ǥ(�ڻ�)
    Const TYPE_2 = 1	' ǥ�ش�������ǥ(��ä�ں�)
    Const TYPE_3 = 2	' �հ�ǥ�ش�������ǥ(�ڻ�)
    Const TYPE_4 = 3	' �հ�ǥ�ش�������ǥ(��ä�ں�)
    Dim arrHTFBody(3)

    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False		' -- ���� �߻����� 
    blnFirst = True			' -- ó�� �����ڵ� ������� 
    blnFile = False			' -- ������������ 
    
    PrintLog "MakeHTF_W1101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1101MA1"

	Set lgcTB_3_2 = New C_TB_3_2		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_3_2.LoadData Then Exit Function
	
	Set cDataExists = new TYPE_DATA_EXIST_W1101MA1			
	
	Call lgcTB_3_2.Clone(oRs2)
	'==========================================
	' -- ��3ȣ��2 ǥ�ش�������ǥ �������� 

	' -- �Ϲݹ��� - �ڻ� 
	If lgcTB_3_2.W1 = "1" Then ' -- �Ϲݹ��� 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_1) = "83" : arrHTFBody(TYPE_3) = "83"
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("A112", 4)		' �Ϲݹ���_�ڻ� �����ڵ� 
		arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNIChar("A172", 4)		' �Ϲݹ���_�ڻ� �����ڵ� (�հ�)
		
		lgcTB_3_2.Filter "W2 LIKE '1%'"					' ���ڵ�� ���� 
		oRs2.Filter  = "W2 LIKE '1%'"	
	
		Do Until lgcTB_3_2.EOF 
		
		'Response.End 'ZZZ
		  SELECt Case  lgcTB_3_2.W4 
		     Case "01"
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '16'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'�ڵ�(01)�����ڻ�= �ڵ� 02 + 16
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(01)�����ڻ�","�ڵ� 02 + 16"))
						   blnError = True	
					End If
					
			  Case "02"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '03'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'�ڵ�(02)�����ڻ�  = �ڵ� 03 + 04 + 05 + 06 + 07 + 10 + 14 + 15	
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(02)�����ڻ�","�ڵ� 03 + 04 + 05 + 06 + 07 + 10 + 14 + 15	"))
						   blnError = True	
					End If		
			Case "07"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '08'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- �ڵ�(07)�ܱ�뿩��  = �ڵ� 08 + 09     
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(07)�ܱ�뿩�� ","�ڵ� 08 + 09 "))
						   blnError = True	
					End If	
			Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(10)�̼���  = �ڵ� 11 + 12 + 13 
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(07)�ܱ�뿩�� ","�ڵ� 08 + 09 "))
						   blnError = True	
					End If	
			  Case "16"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '17'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '29'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'�ڵ�(16)����ڻ�  = �ڵ� 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(16)����ڻ�","�ڵ� 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29"))
						   blnError = True	
					End If			
				Case "36"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '37'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(36)�����ڻ�  = �ڵ� 37 + 50 + 60
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(36)�����ڻ� ","�ڵ� 37 + 50 + 60"))
						   blnError = True	
					End If											
				
				Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '45'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(37)�����ڻ�= �ڵ� 38 + 39 + 40 + 44 + 45 + 46 + 47 + 48 + 49
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(37)�����ڻ�","�ڵ� 38 + 39 + 40 + 44 + 45 + 46 + 47 + 48 + 49"))
						   blnError = True	
					End If		
					
					
		
			 
			    
			    
				Case "50"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '51'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- �ڵ�(50)�����ڻ�= �ڵ� 51 + 52 + 53 + 54 + 55 + 56 + 57 + 58 + 59
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(50)�����ڻ�","�ڵ� 51 + 52 + 53 + 54 + 55 + 56 + 57 + 58 + 59"))
						   blnError = True	
					End If		
					
				Case "60"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '61'"	
					dblAmt1 =UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '68'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '70'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '71'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'-  �ڵ�(60)�����ڻ�= �ڵ� 61 + 62 + 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70 + 71
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(60)�����ڻ�","�ڵ� 61 + 62 + 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70 + 71"))
						   blnError = True	
					End If					
				Case "72"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '36'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					dblAmt72 =UNICDbl(lgcTB_3_2.DR_INV, 0)	
					
					'-  -- �ڵ�(72)�ڻ��Ѱ�= �ڵ� 01 + 36
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(72)�ڻ��Ѱ�","�ڵ� 01 + 36"))
						   blnError = True	
					End If	
					
												
			End Select
			 
			    ' -- �ܾ� 200703 ����?
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.DR_INV, 16, 0)
			
				' -- �հ�(����)	
				If Not ChkNotNull(lgcTB_3_2.W5, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNINumeric(lgcTB_3_2.W5, 20, 0)
				' -- �հ�(�뺯)
				If Not ChkNotNull(lgcTB_3_2.W6, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNINumeric(lgcTB_3_2.W6, 20, 0)
			
			lgcTB_3_2.MoveNext 
		Loop

		lgcTB_3_2.Filter ""			' -- ���� ���� 
		oRs2.Filter = ""
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("", 38)	' -- ���� 
		arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNIChar("", 54)	' -- ���� (�հ�)

		'----------------------------------------------------------------------
		'��ä�� �ں��Ѱ� 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_2) = "83" : arrHTFBody(TYPE_4) = "84"
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("A113", 4)		' �Ϲݹ���_��ä�ں� �����ڵ� 
		arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNIChar("A172", 4)		' �Ϲݹ���_�ڻ� �����ڵ� (�հ�)

		lgcTB_3_2.Filter "W2 LIKE '2%' OR W2 LIKE '3%' OR W2 = '4'"					' ���ڵ�� ���� 
	
		oRs2.Filter = "W2 LIKE '2%' OR W2 LIKE '3%' OR W2 = '4'"					' ���ڵ�� ���� 
		Do Until lgcTB_3_2.EOF 
	
	    	Select Case  lgcTB_3_2.W4 
		        '�ڵ�(01)������ä= �ڵ� 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14
		        Case "01"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '03'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '11'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					''�ڵ�(01)������ä= �ڵ� 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(01)������ä ","�ڵ� 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14"))
						   blnError = True	
					End If	
					
					
					
				 
		        Case "06"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '07'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					''�ڵ�(06)������  = �ڵ� 07 + 08 + 09
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(06)������","�ڵ� 07 + 08 + 09"))
						   blnError = True	
					End If	
						
						
						
				  
		        Case "15"
		         
			        oRs2.MoveFirst
					oRs2.Find "W4 = '16'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					' '- �ڵ�(15)������ä= �ڵ� 16 + 67 + 17 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15)������ä","�ڵ� 16 + 67 + 17 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28"))
						   blnError = True	
					End If	
					
					
					
			  Case "17"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '18'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					'- �ڵ�(17)������Ա�= �ڵ� 18 + 19 + 20
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(17)������Ա�","�ڵ� 18 + 19 + 20"))
						   blnError = True	
					End If	
					
			
				
			    Case "29"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'- �ڵ�(29)��ä�Ѱ�= �ڵ� 01 + 15	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(29)��ä�Ѱ�","�ڵ� 01 + 15"))
						   blnError = True	
					End If	
					
					
				Case "41"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '42'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '43'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'-- �ڵ�(41)�ں���= �ڵ� 42 + 43
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(41)�ں���","�ڵ� 42 + 43"))
						   blnError = True	
					End If	
						
			
		   	   '�ڵ�(41)�� �ں����� 0�� �ƴϸ� �ں��ݰ���������������(��) (A144)��  �ڵ�(01)�ں����� �׸�(5)�⸻�ܾװ� ��ġ 
			
			  
			        If  UNICDbl(lgcTB_3_2.CR_INV, 0) <> 0 Then
			        	Set cDataExists.A144 = new C_TB_50A	' -- W7105MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A144.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A144.WHERE_SQL = " AND W_CD = '01' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A144.LoadData() Then
								blnError = True
						
							Else
							
							     sTmp =  UNICDbl(cDataExists.A144.W5,0)
							    
								If UNICDbl(lgcTB_3_2.CR_INV, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(41)�� �ں���"," �ں��ݰ���������������(��) (A144)��  �ڵ�(01)�ں����� �׸�(5)�⸻�ܾ�"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A144 = Nothing				
					 End If		
			    
			    
			    
			       Case "44"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '45'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'-�ڵ�(44)�ں��׿���= �ڵ� 45 + 46 + 47 + 48 + 49
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(44)�ں��׿���","�ڵ� 45 + 46 + 47 + 48 + 49"))
						   blnError = True	
					End If	
					
					
				 Case "50"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '51'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					'- �ڵ�(50)�����׿���= �ڵ� 51 + 52 + 53 + 54 + 55 + 56
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(50)�����׿���=","�ڵ� 51 + 52 + 53 + 54 + 55 + 56"))
						   blnError = True	
					End If	
	
			     Case "57"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '58'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					'- �ڵ�(57)�ں�����= �ڵ� 58 + 59 + 60 + 61 + 62 + 63 + 64
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(57)�ں�����=","�ڵ� 58 + 59 + 60 + 61 + 62 + 63 + 64"))
						   blnError = True	
					End If	
				
				
				Case "65"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '41'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'	- �ڵ�(65)�ں��Ѱ�= �ڵ� 41 + 44 + 50 + 57
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(65)�ں��Ѱ�=","�ڵ� 41 + 44 + 50 + 57"))
						   blnError = True	
					End If	
				Case "66"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					dblAmt66 =  UNICDbl(lgcTB_3_2.CR_INV, 0)
					'	- �ڵ�(66)��ä�� �ں��Ѱ�= �ڵ� 29 + 65
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(66)��ä�� �ں��Ѱ�","�ڵ� 29 + 65"))
						blnError = True	
					End If		
					
					
			    End Select
			    
			' -- �ܾ� 
			If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.CR_INV, 16, 0)
				
			' -- �հ�(����)	
			If Not ChkNotNull(lgcTB_3_2.W5, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNINumeric(lgcTB_3_2.W5, 20, 0)
			' -- �հ�(�뺯)
			If Not ChkNotNull(lgcTB_3_2.W6, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNINumeric(lgcTB_3_2.W6, 20, 0)
			
			lgcTB_3_2.MoveNext 
		Loop
		
			If dblAmt72 <> dblAmt66  Then
				Call SaveHTFError(lgsPGM_ID, dblAmt72, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(72)�ڻ��Ѱ�(" & dblAmt72 &")","�ڵ�(66)��ä�� �ں��Ѱ�(" & dblAmt66 &")"))
				blnError = True	
		   End If	
		
		
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("", 58)	' -- ���� 
		arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNIChar("", 44)	' -- ���� (�հ�)
	
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_1)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_2)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_3)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_4)
		
		' -- ���Ͽ� ����Ѵ�.
		If Not blnError Then
			Call WriteLine2File(arrHTFBody(TYPE_1))
			Call WriteLine2File(arrHTFBody(TYPE_2))
			Call WriteLine2File(arrHTFBody(TYPE_3))
			Call WriteLine2File(arrHTFBody(TYPE_4))
			
			'Call PushRememberDoc(arrHTFBody(TYPE_3))	' -- �ٷ� ������� �ʰ� ����Ų��(inc_HomeTaxFunc.asp�� ����)
			'Call PushRememberDoc(arrHTFBody(TYPE_4))
		End If		

	
	Else	' -- �������� 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_1) = "83" : arrHTFBody(TYPE_2) = "85"
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("A114", 4)		' �������� �����ڵ� 
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("A172", 4)		' �������� �����ڵ� 
		
		Do Until lgcTB_3_2.EOF 
		
			If Left(lgcTB_3_2.W1, 1) = "1" Then
			
			   Select Case lgcTB_3_2.W4 	
				Case "01"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- �ڵ�(01)���ݰ���ġ��= �ڵ� 02 + 03 + 04
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(01)���ݰ���ġ��","�ڵ� 02 + 03 + 04"))
						blnError = True	
					End If	
				Case "05"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '06'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- �ڵ�(05)��ǰ��������= �ڵ� 06 + 07 + 08 + 09
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(05)��ǰ��������","�ڵ� 06 + 07 + 08 + 09"))
						blnError = True	
					End If				
					
				Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- �ڵ�(10)������������= �ڵ� 11 + 12 + 13 + 14
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(10)������������","�ڵ� 11 + 12 + 13 + 14"))
						blnError = True	
					End If			
				Case "15"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '16'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(15)����ä��= �ڵ� 16 + 17 + 18 + 19 + 20 + 21
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(15)����ä��","�ڵ� 16 + 17 + 18 + 19 + 20 + 21"))
						blnError = True	
					End If
				Case "23"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '24'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(23)�����ڻ�= �ڵ� 24 + 25 + 26 + 27
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(23)�����ڻ�","�ڵ� 24 + 25 + 26 + 27"))
						blnError = True	
					End If	
				Case "28"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '30'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '35'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '36'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - �ڵ�(28)�����ڻ�= �ڵ� 29 + 30 + 35 + 36
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(28)�����ڻ�"," �ڵ� 29 + 30 + 35 + 36"))
						blnError = True	
					End If	
				Case "30"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '31'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '32'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '33'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '34'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- �ڵ�(30)�����ڻ�= �ڵ� 31 + 32 + 33 + 34
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(30)�����ڻ�"," �ڵ� 31 + 32 + 33 + 34"))
						blnError = True	
					End If		
																			
				Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '41'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '42'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- - �ڵ�(37)��Ÿ�ڻ�= �ڵ� 38 + 39 + 40 + 41 + 42
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(37)��Ÿ�ڻ�"," �ڵ� 38 + 39 + 40 + 41 + 42"))
						blnError = True	
					End If			
				
				Case "43"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '37'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					dblAmt43 =  UNICDbl(lgcTB_3_2.DR_INV, 0)
					'�ڵ�(43)�ڻ��Ѱ�= �ڵ� 01 + 05 + 10 + 15 + 22 + 23 + 28 + 37
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(43)�ڻ��Ѱ�"," �ڵ�01 + 05 + 10 + 15 + 22 + 23 + 28 + 37"))
						blnError = True	
					End If				
				 End Select
				
			
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.W3, DR_INV, 0)
				
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.W3, DR_INV, 0)
			Else
			
			   Select Case lgcTB_3_2.W4 	
			   	Case "44"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '45'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					'- �ڵ�(44)�ſ��ä= �ڵ� 45 + 46
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(44)�ſ��ä"," �ڵ� 45 + 46"))
						blnError = True	
					End If		
																											
				
				
				Case "47"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '48'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '51'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
										
					'- �ڵ�(47)���Ա�= �ڵ� 48 + 49 + 50 + 51 + 52	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(47)���Ա�"," �ڵ� 48 + 49 + 50 + 51 + 52"))
						blnError = True	
					End If																							
			
				Case "54"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '55'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
										
					'- �ڵ�(54)��Ÿ��ä= �ڵ� 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(54)��Ÿ��ä"," �ڵ�  55 + 56 + 57 + 58 + 59 + 60 + 61 + 62"))
						blnError = True	
					End If		
					
					
					
				Case "63"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '64'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- �ڵ�(63)����� ���غ��= �ڵ� 64 + 65 + 66	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(63)����� ���غ��"," �ڵ� 64 + 65 + 66"))
						blnError = True	
					End If					
				
				Case "67"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '44'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- �ڵ�(67)��ä�Ѱ�= �ڵ� 44 + 47 + 53 + 54 + 63
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(67)��ä�Ѱ�"," �ڵ� 44 + 47 + 53 + 54 + 63"))
						blnError = True	
					End If		
				
				Case "68"
			     
						
			
		   	   '�ڵ�(68)�� �ں����� 0�� �ƴϸ� �ں��ݰ���������������(��) (A144)��  �ڵ�(01)�ں����� �׸�(5)�⸻�ܾװ� ��ġ 
			
			  
			        If  UNICDbl(lgcTB_3_2.CR_INV, 0) <> 0 Then
			        	Set cDataExists.A144 = new C_TB_50A	' -- W7105MA1_HTF.asp �� ���ǵ� 
								
							' -- �߰� ��ȸ������ �о�´�.
							cDataExists.A144.CALLED_OUT	= True		' -- �ܺο��� ȣ������ �˸� 
							cDataExists.A144.WHERE_SQL = " AND W_CD = '01' "			' -- Ŭ������ �⺻������ �ε��ϴ� ���ǿ��� �߰� ������ ������ 
								
							If Not cDataExists.A144.LoadData() Then
								blnError = True
						
							Else
							
							     sTmp =  UNICDbl(cDataExists.A144.W5,0)
							    
								If UNICDbl(lgcTB_3_2.CR_INV, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(68)�� �ں���"," �ں��ݰ���������������(��) (A144)��  �ڵ�(01)�ں����� �׸�(5)�⸻�ܾ�"))
								End If
							End If
					
							
							' -- ����� Ŭ���� �޸� ���� 
							Set cDataExists.A144 = Nothing				
					 End If		
			    
			    
			    	
				
				
				Case "69"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '70'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '71'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '72'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- �ڵ�(69)�ں��׿���= �ڵ� 70 + 71 + 72
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(69)�ں��׿���"," �ڵ� 70 + 71 + 72"))
						blnError = True	
					End If		
				
				Case "73"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '74'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '75'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '76'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '77'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '78'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- - �ڵ�(73)�����׿���= �ڵ� 74 + 75 + 76 + 77 + 78
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(73)�����׿���"," �ڵ� 74 + 75 + 76 + 77 + 78"))
						blnError = True	
					End If	
				
				Case "79"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '80'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '81'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '82'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)

															
					' �ڵ�(79)�ں�����= �ڵ� 80 + 81 + 82	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(79)�ں�����"," �ڵ�  80 + 81 + 82	"))
						blnError = True	
					End If		
						
				Case "83"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '68'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '73'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '79'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					' - �ڵ�(83)�ں��Ѱ�= �ڵ� 68 + 69 + 73 + 79
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(83)�ں��Ѱ�"," �ڵ�  68 + 69 + 73 + 79"))
						blnError = True	
					End If				
				
				Case "84"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '67'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '83'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
															
					' - �ڵ�(84)��ä �� �ں��Ѱ�= �ڵ� 67 + 83
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(84)��ä �� �ں��Ѱ�"," �ڵ�  67 + 83"))
						blnError = True	
					End If		
				
				Case "84"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '67'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '83'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					dblAmt84 = UNICDbl(lgcTB_3_2.CR_INV, 0)
															
					' - �ڵ�(84)��ä �� �ں��Ѱ�= �ڵ� 67 + 83
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(84)��ä �� �ں��Ѱ�"," �ڵ�  67 + 83"))
						blnError = True	
					End If	
						
			    End Select
			    
				If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.W3, CR_INV, 0)
				
				If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.W3, CR_INV, 0)
			End If	
		
			lgcTB_3_2.MoveNext 
		Loop
		
		If dblAmt43 <> dblAmt84  Then
			Call SaveHTFError(lgsPGM_ID, dblAmt43, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�ڵ�(43)�ڻ��Ѱ�(" & dblAmt43 &")","�ڵ�(84)��ä�� �ں��Ѱ�(" & dblAmt84 &")"))
			blnError = True	
		 End If
		
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("", 50)	' -- ���� 
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("", 134)	' -- ���� 
	
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_1)
		
		' -- ���Ͽ� ����Ѵ�.
		If Not blnError Then
			Call WriteLine2File(arrHTFBody(TYPE_1))
			Call WriteLine2File(arrHTFBody(TYPE_2))
			'Call PushRememberDoc(arrHTFBody(TYPE_2))	' -- �ٷ� ������� �ʰ� ����Ų��(inc_HomeTaxFunc.asp�� ����)
		End If
	End If
					
	' ----------- 
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_3_2 = Nothing	' -- �޸����� 
	
End Function


%>
