
<% session.CodePage=949 %>

<%
' -------- Ȩ�ؽ� ���� �Լ� -------------

' -------- �޽��� -----------
Const TYPE_DATA_NOT_FOUND	= "�ش� ���� ����Ÿ�� �������� �ʽ��ϴ�."
Const TYPE_CHK_NULL			= "%1 ��(��) �ʼ��Է� �׸��Դϴ�"
Const TYPE_CHK_DATE			= "%1 ��(��) ��¥���� �����Դϴ�"
Const TYPE_CHK_NUM			= "%1 ��(��) �������� �����Դϴ�"
Const TYPE_CHK_ADDR			= "%1 ��(��) ����,Ư������ �������Դϴ�"
Const TYPE_CHK_TEL_NO		= "%1 ��(��) ��ȭ��ȣ ���Ŀ����Դϴ�"
Const TYPE_SYSTEM_ERROR		= "�����߻�: %1"
Const TYPE_FILE_NOT_MAKE	= "%1 ��ο� ������ ������ �� �����ϴ�"
Const TYPE_CHK_DATE_FRTO	= "%1 �� %2 �Ϻ��� ������ �� �����ϴ�"
Const TYPE_CHK_DATE_OVER	= "������� �Ⱓ�� 1���� �ʰ��� �� �����ϴ�"
Const TYPE_CHK_BOUNDARY		= "%1�ڵ�� BOUNDARY(%2)�� �������� �ʽ��ϴ�"
Const TYPE_CHK_RGST_NO		= "����ڹ�ȣ(4:2)��"
Const TYPE_CHK_ZERO_OVER	= "%1 ��(��) 0���� Ŀ�� �մϴ�."
Const TYPE_CHK_NOT_EQUAL	= "%1 ��(��) %2 �� ��ġ�����ʽ��ϴ�"
Const TYPE_CHK_LOW_AMT		= "%1 �� %2 ���� �۽��ϴ�"
Const TYPE_CHK_HIGH_AMT		= "%1 �� %2 ���� Ů�ϴ�"
Const TYPE_CHK_HTF_MODULE	= "%1 ���ڽŰ� ����� �ۼ����� �ʾҽ��ϴ�"
Const TYPE_MSG_NORMAL_ERR	= "%1 ��(��) �����Դϴ�"
Const TYPE_CHK_LOW_EQUAL_AMT	= "%1�� %2 ���� �۰ų� ���ƾ��մϴ�"
Const TYPE_CHK_ZERO_EQUAL	= "%1 ��(��) 0�̾���մϴ�"
Const TYPE_CHK_OVER_EQUAL	= "%1�� %2 ���� ũ�ų� ���ƾ��մϴ�"
Const TYPE_CHK_CHARNUM		= "%1�� �ڸ�����  %2 �̾���մϴ�"
Const TYPE_POSITIVE			= "%1�� ����̾���մϴ�."


Function UNIGetMesg(Byval pTemp, Byval pMsg1, Byval pMsg2)
	If pMsg1 = "" Then
		UNIGetMesg = pTemp
	ElseIf pMsg1 <> "" And pMsg2 <> "" Then
		UNIGetMesg = Replace(Replace(pTemp, "%1", pMsg1), "%2", pMsg2)
	ElseIf pMsg1 <> "" And pMsg2 = "" Then
		UNIGetMesg = Replace(pTemp, "%1", pMsg1)
	End If
End Function

' -------- ���ڿ� ���� ��ƿ��Ƽ -----------
Function LeftH(ByVal strString, ByVal lngLength)
    Dim i, iLen, sTmp, iLen2, sRet
        
    For i = 1 To lngLength
        sTmp = Mid(strString, i, 1)
        sRet = sRet & sTmp
        
        If Asc(sTmp) < 0 Then    ' �ѱ� 
            iLen2 = iLen2 + 2
        Else
            iLen2 = iLen2 + 1
        End If
        If iLen2 = lngLength Then Exit For
    Next
    'LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    LeftH = sRet
End Function
' -------- ����/���� ��ȯ -----------
' -- CHAR���ڷ� ���� 
Function UNIChar(Byval pData, Byval pLen)
	Dim pTmp
	pTmp = String(pLen, " ")
	pTmp = LeftH(pData & pTmp, pLen)
	UNIChar = pTmp
End Function


Function UNINumeric(Byval pData, Byval pLen, Byval pDec)
	Dim pTmp, pRet , pTmp2,  pRet2 , pMesg, pArr
	
	Err.Clear 
	'On Error Resume Next
	
	If Instr(1, CStr(pData), ".") > 0 Then
		pArr = Split(CStr(pData), ".")
		pTmp = pArr(0)
		pTmp2= pArr(1)
	Else
		pTmp = pData
		pTmp = 0
	End If

	If Err Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_NUM, pData, ""))
		Exit Function
	End If

	pRet =  String(pLen, "0")
	pRet2 = String(pDec, "0")

	If unicdbl(pData,0) >= 0  then
	   pRet = Right(pRet & pTmp, pLen - pDec) & Right(pRet2 & pTmp2, pDec)
	Else
	   pTmp = Replace(pTmp, "-", "")  '�����ϰ�� �� ���ڸ��� ǥ�� 
	   pRet= "-" & Right(pRet & pTmp, (pLen-pDec)-1) & Right(pRet2 & pTmp2, pDec)
	End if  

	UNINumeric = pRet
End Function


' -- NUMERIC ���� 
Function UNINumeric_old(Byval pData, Byval pLen, Byval pDec)
	Dim pTmp, pRet , pTmp2,  pRet2 , pMesg
	
	Err.Clear 
	On Error Resume Next
	
	   pTmp  = Replace(fix(unicdbl(pData,0)), ".", "")   ' �����κ� 
	   pTmp2 = Replace(unicdbl(pData,0) - pTmp, ".", "")  '�Ҽ��κ� 
	   
   
	
	If Err Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_NUM, pData, ""))
		Exit Function
	End If
	
	  pRet =  String(pLen, "0")
	  pRet2 = String(pDec, "0")
	 If unicdbl(pData,0) >= 0  then
	  
	   pRet = Right(pRet & pTmp, pLen - pDec) & Right(pTmp & pRet2, pDec)
	Else
	   pTmp = Replace(pTmp, "-", "")  '�����ϰ�� �� ���ڸ��� ǥ�� 
	   pRet= "-" & Right(pRet & pTmp, (pLen-pDec)-1) & Right(pTmp & pRet2, pDec)
	End if  
	
	UNINumeric = pRet
End Function

Const TYPE_NOT_NULL = 0
Const TYPE_NULL		= 1
Const TYPE_NOT_NULL_DEFAULT_0 = 2

' -------- �ΰ� üũ(����,����,��¥) -----------
Function ChkNotNull(Byval pData, Byval pMesg)
	ChkNotNull = False
	If IsNull(pData) Or Trim(pData) = "" Then
		Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg(TYPE_CHK_NULL, pMesg,""))
		Exit Function
	End If
	ChkNotNull = True
End Function

' -- ��¥ ���� üũ 
Function ChkDate(Byval pDate, Byval pMesg)
	ChkDate = False
	If Not IsNull(pDate) Then
		If Not IsDate(pDate) Then
			Call SaveHTFError(lgsPGM_ID, pDate, UNIGetMesg(TYPE_CHK_DATE, pMesg,""))
			Exit Function
		End If
	End If
	ChkDate = True
End Function

' -- ���ڰ˻� 
Function ChkNumeric(Byval pNum, Byval pMesg)
	ChkNumeric = False
	If Not IsNumeric(pNum) Then
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_NUM, pMesg,""))
		Exit Function
	End If
	ChkNumeric = True
End Function

Function ChkMinusAmt(Byval pNum, Byval pMesg)
	ChkMinusAmt = False
	If UNICDbl(pNum, 0) < 0 Then
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_ZERO_OVER, pMesg,""))
		Exit Function
	End If
	ChkMinusAmt = True
End Function

' -- 8�ڸ� ��¥ ���� 
Function UNI8Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI8Date = String(8, " ")
		Exit Function
	End If
	UNI8Date = Year(pDate) & Right("0" & Month(pDate), 2) & Right("0" & Day(pDate), 2)
End Function

' -- 6�ڸ� ��¥ ���� 
Function UNI6Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI6Date = String(6, " ")
		Exit Function
	End If
	UNI6Date = Year(pDate) & Right("0" & Month(pDate), 2) 
End Function

' -- 4�ڸ� ��¥ ���� 
Function UNI4Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI4Date = String(4, " ")
		Exit Function
	End If
	UNI4Date = Year(pDate)
End Function



' -- ����/���� ��¥ �� 
Function ChkDateFrTo(Byval pFrDt, Byval pToDt, Byval pMesg1, Byval pMesg2, Byval blnOver)
	ChkDateFrTo = False
	If pFrDt > pToDt Then	' -- �������� �����Ϻ��� ũ�� ���� 
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_DATE_FRTO, pMesg1, pMesg2))
		Exit Function
	ElseIf DateDiff("m", pFrDt, pToDt) > 12 And blnOver = False Then	' -- 12�������� ũ�� 
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_DATE_OVER, "", ""))
		Exit Function
	End If	
	ChkDateFrTo = True
End Function

' -- Dash ����(�ֹε�Ϲ�ȣ��)
Function UNIRemoveDash(Byval pData)
	UNIRemoveDash = Replace(pData, "-", "")
End Function

' -- ������ �������� �ʴ°� �˻� 
Function ChkContents(ByVal pData, Byval pMesg)
    Dim i, iLen, sTmp, blnError
	On Error Resume Next
    blnError = False
    iLen = Len(pData)
    For i = 1 To iLen
        sTmp = Asc(Mid(pData, i, 1))
        If sTmp >= 65 And sTmp <= 90 Then    ' �ѱ��� ����/����� ����� ���� 
            blnError = True
            Exit For   ' -- �������Խ�Ʈ���� 
        ElseIf sTmp >= 97 And sTmp <= 122 Then
            blnError = True
            Exit For   ' -- �������Խ�Ʈ���� 
        ElseIf sTmp = 32 Then
            If i < iLen And Asc(Mid(pData, i + 1, 1)) = 32 Then ' -- ������ 1���϶��� ��� 
                blnError = True
				Exit For   ' -- ���� 2�� 
            End If
        ElseIf sTmp > 0 And sTmp < 32 Then
            blnError = True
            Exit For   ' -- Ư������ 
        ElseIf sTmp > 32 And sTmp < 45 Then     ' - �뽬(45)�� ��� 
            blnError = True
            Exit For   ' -- Ư������ 
        ElseIf sTmp > 45 And sTmp < 48 Then
            blnError = True
            Exit For   ' -- Ư������ 
        ElseIf sTmp > 57 And sTmp < 65 Then
            blnError = True
            Exit For   ' -- Ư������ 
        End If
    Next
 
    If blnError = True  Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_ADDR, pMesg,""))
	ElseIf Err.number > 0 Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_SYSTEM_ERROR, Err.Description ,""))
    End If
    ChkContents = Not blnError
End Function

Function ChkTelNo(Byval pData, Byval pMesg)
    Dim i, iLen, sTmp, blnError
    ChkTelNo = False
    sTmp = Replace(pData, "-", "")
	If sTmp <> "" And (Not isNumeric(sTmp) Or Left(pData, 1) <> "0") Then 
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_TEL_NO, pMesg,""))
		Exit Function
    End If
    ChkTelNo = True
End Function

' -- �Ϸù�ȣ ����: 000001, 000002
Function UNISeqNo6(Byval pData)
	Dim pTmp
	pTmp = String(6, "0") & pData
	UNISeqNo6 = Right(pTmp, 6)
End Function

Function UNISeqNo(Byval pData, Byval pLen)
	Dim pTmp
	pTmp = String(pLen, "0") & pData
	UNISeqNo = Right(pTmp, pLen)
End Function

' -- �ٿ���� üũ 
Function ChkBoundary(Byval pBoundary, Byval pData, Byval pMesg)
	ChkBoundary = False
	If IsNull(pData) Or Trim(pData) = "" Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_NULL, pMesg,""))
		Exit Function
	ElseIf Instr(1, pBoundary, pData) = 0 Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_BOUNDARY, pMesg, pBoundary))
		Exit Function
	End If
	ChkBoundary = True
End Function

' -- ���� ����ڵ�Ϲ�ȣ���� 4:2 ����Ÿ 
Function GetRgstNo42(Byval pRgstNo)
	GetRgstNo42 = Mid(Replace(pRgstNo, "-", ""), 4, 2)
End Function


' -- �ڸ��� üũ(������, �ڸ���)
Function ChkCharNum(Byval pData , Byval pNum , Byval pMesg)
Dim pTmp
     ChkCharNum = False
	  pTmp = Len(Replace(pData, "-", ""))
	  if unicdbl(pTmp,0) <> unicdbl(pNum,0) then
	     Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_CHARNUM, pMesg, pNum))
		 Exit Function
	     
	  End if 
	 ChkCharNum = True
End Function


' ---------------------- Ȩ�ؽ� ���� �����Լ� --------------
Dim lgFSO , lgStream , lgobj

Set lgFSO = Nothing
Set lgStream = Nothing
Set lgobj = Nothing
Function InitFileSystem(Byval pPath, Byval pFileNM)
	Dim sFullFileNm
	Err.Clear 
	On Error Resume Next
	sFullFileNm =  Server.MapPath(pPath) & "\" & pFileNM
	PrintLog "InitFileSystem .. : " & sFullFileNm

	Set lgFSO = Server.CreateObject("Scripting.FileSystemObject")
	

	if not lgFSO.FolderExists(Server.MapPath(pPath)) then
	     set lgobj = lgFSO.CreateFolder(Server.MapPath(pPath))
	end if

	Set lgStream = lgFSO.CreateTextFile(sFullFileNm, true)
	
	If Err.number > 0 Then
		PrintLog "InitFileSystem Error.. : " & Err.Description
			
		Call SaveHTFError(lgsPGM_ID, sFullFileNm, UNIGetMesg(TYPE_FILE_NOT_MAKE, sFullFileNm,""))
		Response.End
	End If
End Function

Function WriteLine2File(Byval pDesc)
	On Error Resume Next
	If Not lgStream is Nothing Then
		If lgblnPushDoc Then	' -- ����� ���� ����� 
		
			' ��¹���(pDesc)�� ����ȹ���(lgarrTAX_DOC)�� �����ڵ带 ���� ����� ������ �����ڵ尡 ������ ����Ѵ�.
			Dim pDoc1, pDoc2, i, iMaxCnt
			
			pDoc1 = Mid(pDesc, 4, 3)	' 4��°���� 3�ڸ��б� (��: 83A101....  101�� ����)
			iMaxCnt = UBound(lgarrTAX_DOC)
			
			For i = 0 To iMaxCnt
				If Trim(lgarrTAX_DOC(i)) <> "" Then
				
					pDoc2 = Mid(lgarrTAX_DOC(i), 4, 3)	' 4��°���� 3�ڸ��б� 
					If CDbl(pDoc1) > CDbl(pDoc2) Then
						If Err Then
							PrintLog "WriteLine2File lgarrTAX_DOC Error.. : " & Err.Description
							PrintLog "pDoc1=" & pDoc1
							PrintLog "pDoc2=" & pDoc2
							
							Exit Function
						End If
						' -- ����ؾߵ� ������ �ڵ尡 ���� ������ �ڵ庸�� ũ�ٸ�,  ���� ���� ���� ����Ѵ�.
						'PrintLog "lgarrTAX_DOC(" & i & ")=" & lgarrTAX_DOC(i)
						lgStream.WriteLine lgarrTAX_DOC(i)
						lgarrTAX_DOC(i) = ""
					End If
				End If
			Next
			
		End If
		
		lgStream.WriteLine pDesc		' -- ��û�� ������ ����Ѵ�.
	End If
End Function

Function Write2File(Byval pDesc)
	On Error Resume Next
	If Not lgStream is Nothing Then
		If lgblnPushDoc Then	' -- ����� ���� ����� 
		
			' ��¹���(pDesc)�� ����ȹ���(lgarrTAX_DOC)�� �����ڵ带 ���� ����� ������ �����ڵ尡 ������ ����Ѵ�.
			Dim pDoc1, pDoc2, i, iMaxCnt
			
			pDoc1 = Mid(pDesc, 4, 3)	' 4��°���� 3�ڸ��б� (��: 83A101....  101�� ����)
			iMaxCnt = UBound(lgarrTAX_DOC)
			
			For i = 0 To iMaxCnt
				If Trim(lgarrTAX_DOC(i)) <> "" Then
				
					pDoc2 = Mid(lgarrTAX_DOC(i), 4, 3)	' 4��°���� 3�ڸ��б� 
					If uniCDbl(pDoc1,0) > UniCDbl(pDoc2,0) Then
						If Err Then
							PrintLog "Write2File lgarrTAX_DOC Error.. : " & Err.Description
							PrintLog "pDoc1=" & pDoc1
							PrintLog "pDoc2=" & pDoc2
							
							Exit Function
						End If
						' -- ����ؾߵ� ������ �ڵ尡 ���� ������ �ڵ庸�� ũ�ٸ�,  ���� ���� ���� ����Ѵ�.
						'PrintLog "lgarrTAX_DOC(" & i & ")=" & lgarrTAX_DOC(i)
						
						Select Case Mid(lgarrTAX_DOC(i), 3, 4)
							Case "A172"
								lgStream.Write lgarrTAX_DOC(i) & vbCrLf
							Case Else
								lgStream.Write lgarrTAX_DOC(i)
						End Select
						lgarrTAX_DOC(i) = ""
					End If
				End If
			Next
			
		End If
		
		lgStream.Write pDesc		' -- ��û�� ������ ����Ѵ�.
	End If
End Function

Function CloseFileSystem()

	' -- ���� ������ �����ִ°� �ִٸ� ����� �����Ѵ� 
	Dim iMaxCnt, i
	
	If lgblnPushDoc Then	' -- ����� ���� ����� 

		iMaxCnt = UBound(lgarrTAX_DOC)
				
		For i = 0 To iMaxCnt
			If Trim(lgarrTAX_DOC(i)) <> "" Then
				PrintLog "���� ������ ������ ����մϴ� : " & lgarrTAX_DOC(i)	
				' -- ����ؾߵ� ������ �ڵ尡 ���� ������ �ڵ庸�� ũ�ٸ�,  ���� ���� ���� ����Ѵ�.  :** ���� ������ ������ �ٸ����� ������������ ��� �Ф� 
				'PrintLog "lgarrTAX_DOC(" & i & ")=" & lgarrTAX_DOC(i)
							
				Select Case Mid(lgarrTAX_DOC(i), 3, 4)
					'Case "A172"	' -- ���Ŀ��� ����Ű�� ���ٶ�..
					'	lgStream.Write lgarrTAX_DOC(i) & vbCrLf
					Case Else
						lgStream.Write lgarrTAX_DOC(i)
				End Select
				lgarrTAX_DOC(i) = ""
			End If
		Next
	End If
	
	If Not lgStream is Nothing Then
		lgStream.Close
	End If
	Set lgStream = Nothing
	Set lgFSO = Nothing
End Function

Dim lgarrTAX_DOC(10)	' -- ���ϼ������ �������� �ʴ� ���ĵ� 
Dim lgblnPushDoc
lgblnPushDoc =False		' -- ����� ���Ͼ��� 

Function PushRememberDoc(Byval pDoc)
	Dim Index
	Index = GetPushIndex()
	If Index = -1 Then
		Err.Raise 60000, "inc_HomeTaxFunc.asp_PushRememberDoc()", "lgarrTAX_DOC �迭 �ʰ�"
	End If
	lgarrTAX_DOC(Index) = pDoc
	PrintLog "PushRememberDoc(" & Index & ")=" & pDoc
	lgblnPushDoc		= True	' -- ����� ���� ���� 
End Function

Function GetPushIndex()	' -- ���� ��� �ִ� �迭�ε��� ���� 
	Dim i , iMaxCnt
	iMaxCnt = UBound(lgarrTAX_DOC)
	For i = 0 To iMaxCnt
		If Trim(lgarrTAX_DOC(i)) = "" Then
			PrintLog "GetPushIndex=" & i
			GetPushIndex = i
			Exit Function
		End If
	Next
	GetPushIndex = -1
End Function

%>