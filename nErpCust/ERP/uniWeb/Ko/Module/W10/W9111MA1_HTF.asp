<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : ��54ȣ �ֽĵ� ������Ȳ����(��)
'*  3. Program ID           : W9111MA1
'*  4. Program Name         : W9111MA1_HTF.asp
'*  5. Program Desc         : ���ڽŰ� Conversion ���α׷� 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

' ------------------ ���� ���� --------------------------------
Dim lgcTB_54_BP

Set lgcTB_54_BP = Nothing ' -- �ʱ�ȭ 

Class C_TB_54_BP
	' -- ���̺��� �÷����� 

	Dim WHERE_SQL		' -- �⺻ �˻�����(����/�������/�Ű���)���� �˻����� 
	Dim	CALLED_OUT		' -- �ܺο��� �θ� ��� 
	
	Private lgoRs1		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs2		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	Private lgoRs3		' -- ��Ƽ�ο� ����Ÿ�� ���������� �����Ѵ�.
	
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
			Case 2
				lgoRs2.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
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
	
	Function RecordCount(Byval pType)
		Select Case pType
			Case 2
				RecordCount = lgoRs2.RecordCount
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
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
	            lgStrSQL = lgStrSQL & " B.TAX_OFFICE, B.FISC_START_DT, B.FISC_END_DT " & vbCrLf
	            lgStrSQL = lgStrSQL & " , B.OWN_RGST_NO, B.CO_NM, B.REPRE_NM " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54_BPH	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & "		INNER JOIN TB_COMPANY_HISTORY	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf	' ����3ȣ 

				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54_BPD	A  WITH (NOLOCK) " & vbCrLf	' ����3ȣ 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- ��ȸ���� �߰� 

		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- ����Ÿ ���� üũ 
Class TYPE_DATA_EXIST_W9111MA1
	Dim A103

End Class

' ------------------ ���� �Լ� --------------------------------
Function MakeHTF_W9111MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, arrTmp1(1), arrTmp2(1)
    Dim iSeqNo, iDx, iDtlCnt
    
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9111MA1"

	Set lgcTB_54_BP = New C_TB_54_BP		' -- �ش缭�� Ŭ���� 
	
	If Not lgcTB_54_BP.LoadData Then Exit Function			' -- ��3ȣ2(1)(2)ǥ�ؼ��Ͱ�꼭 ���� �ε� 
	
	' -- ������ ���� ���� 
	Set cDataExists = new TYPE_DATA_EXIST_W9111MA1

	'==========================================
	' -- ��54ȣ��ǥ�ֽ��������о絵���� �������� 
	' -- 1. �⺻����(83A132)
	sHTFBody = "B"
	sHTFBody = sHTFBody & UNIChar("41", 2)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "TAX_OFFICE"), "�������ڵ�") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "TAX_OFFICE"), 3)
				
	If  ChkNotNull(lgcTB_54_BP.GetData(1, "OWN_RGST_NO"), "�Ű����_����ڵ�Ϲ�ȣ") Then 
	    IF UNIRemoveDash(lgcTB_54_BP.GetData(1, "OWN_RGST_NO")) <>  UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO) Then
	       '�Ű������ ����ڵ�Ϲ�ȣ	- ���μ��Ű����(A100)�� ������ID�� ��ġ 
	        Call SaveHTFError(lgsPGM_ID, lgcTB_54_BP.GetData(1, "OWN_RGST_NO"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "�Ű������ ����ڵ�Ϲ�ȣ", "���μ��Ű����(A100)�� ������ID"))
	        blnError = True	
	    End If 
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(1, "OWN_RGST_NO")), 10)
		
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "CO_NM"), "�Ű����_���θ�") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "CO_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "REPRE_NM"), "�Ű����_��ǥ�ڼ���") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "REPRE_NM"), 30)
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "FISC_START_DT"), "�������(��������)") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(1, "FISC_START_DT"))
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "FISC_END_DT"), "�������(��������)") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(1, "FISC_END_DT"))
		
	If Not ChkBoundary("1,2,3,4,5,6,7", lgcTB_54_BP.GetData(1, "W4"), "�ֽı���") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "W4"), 1)
	
	sHTFBody = sHTFBody & UNINumeric(UNICDBL(lgcTB_54_BP.RecordCount(2),0), 6, 0)	' �׸��� ���� 
		
	sHTFBody = sHTFBody & UNIChar("", 91)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
		
	' -- �ֽļ�������Ȳ 
	iSeqNo = 1	
	
	Do Until lgcTB_54_BP.EOF(2) 

		sHTFBody = sHTFBody & "C"
		sHTFBody = sHTFBody & UNIChar("41", 2)		' Ư���� ��ȭ�� ���ٸ� ȣ�����α׷����� ������ �����ڵ带 ��� 
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TAX_OFFICE, 3)	' -- �������ڵ� 
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO), 10) ' -- ����ڹ�ȣ 
	    
	    
	    If iSeqNo <> 1 Then
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W6"), "�絵��_�ֹ�(�����)��Ϲ�ȣ") Then blnError = True	
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W5"), "�絵��_����(���θ�)") Then blnError = True	
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W9"), "�絵�ֽļ�(�����¼�)") Then blnError = True	
		 End If  
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(2, "W6")), 13)
		'�絵��_�ֹ�(�����)��Ϲ�ȣ	- �絵��_�ֹ�(�����)��Ϲ�ȣ�� �����Ǵ� ��'AAAAAAAAAAAAA'���� ���� �Ϸù�ȣ�� ��000001�� �� �ƴϸ� ���� 
		' �ֹι�ȣ�� ��BBBBBBBBBBBBB���� �ƴϰų� ��CCCCCCCCCC��,�� DDDDDDDDDD��,���� EEEEEEEEEE��,�� FFFFFFFFFF������ �������� �ʰ� �Ǵ� ���� 4�ڸ��� 0000�̰ų� 4�ڸ� �����̸� ���� 
		
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(2, "W5")), 40)
		
		If lgcTB_54_BP.GetData(2, "W7") <> "" Then 
		   sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(2, "W7"))
		Else
		   sHTFBody = sHTFBody & UNIChar( "", 8) 
		End IF   
		
		
		If  lgcTB_54_BP.GetData(2, "W8") <> "" Then 
		   sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(2, "W8"))
		Else
		   sHTFBody = sHTFBody & UNIChar("", 8) 
		End IF   
		   
		
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54_BP.GetData(2, "W9"), 13, 0)	
		
		sHTFBody = sHTFBody & UNIChar("", 96) & vbCrLf

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_54_BP.MoveNext(2)	' -- 1�� ���ڵ�� 
	Loop

	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- ���Ͽ� ����Ѵ�.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- �޸����� 
	Set lgcTB_54_BP = Nothing	' -- �޸����� 
	
End Function

' ------------------ ��ȸ �Լ� --------------------------------
Sub SubMakeSQLStatements_W9111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- �ܺ� ���� SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9111MA1 : " & lgStrSQL
End Sub
%>
