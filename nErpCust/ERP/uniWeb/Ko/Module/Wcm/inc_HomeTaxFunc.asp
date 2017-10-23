
<% session.CodePage=949 %>

<%
' -------- 홈텍스 관련 함수 -------------

' -------- 메시지 -----------
Const TYPE_DATA_NOT_FOUND	= "해당 서식 데이타가 존재하지 않습니다."
Const TYPE_CHK_NULL			= "%1 은(는) 필수입력 항목입니다"
Const TYPE_CHK_DATE			= "%1 은(는) 날짜형식 오류입니다"
Const TYPE_CHK_NUM			= "%1 은(는) 숫자형식 오류입니다"
Const TYPE_CHK_ADDR			= "%1 은(는) 영어,특수문자 사용오류입니다"
Const TYPE_CHK_TEL_NO		= "%1 은(는) 전화번호 형식오류입니다"
Const TYPE_SYSTEM_ERROR		= "에러발생: %1"
Const TYPE_FILE_NOT_MAKE	= "%1 경로에 파일을 생성할 수 없습니다"
Const TYPE_CHK_DATE_FRTO	= "%1 는 %2 일보다 이후일 수 없습니다"
Const TYPE_CHK_DATE_OVER	= "사업연도 기간은 1년을 초과할 수 없습니다"
Const TYPE_CHK_BOUNDARY		= "%1코드는 BOUNDARY(%2)에 존재하지 않습니다"
Const TYPE_CHK_RGST_NO		= "사업자번호(4:2)가"
Const TYPE_CHK_ZERO_OVER	= "%1 은(는) 0보다 커야 합니다."
Const TYPE_CHK_NOT_EQUAL	= "%1 과(와) %2 이 일치하지않습니다"
Const TYPE_CHK_LOW_AMT		= "%1 이 %2 보다 작습니다"
Const TYPE_CHK_HIGH_AMT		= "%1 이 %2 보다 큽니다"
Const TYPE_CHK_HTF_MODULE	= "%1 전자신고 모듈이 작성되지 않았습니다"
Const TYPE_MSG_NORMAL_ERR	= "%1 은(는) 오류입니다"
Const TYPE_CHK_LOW_EQUAL_AMT	= "%1이 %2 보다 작거나 같아야합니다"
Const TYPE_CHK_ZERO_EQUAL	= "%1 은(는) 0이어야합니다"
Const TYPE_CHK_OVER_EQUAL	= "%1이 %2 보다 크거나 같아야합니다"
Const TYPE_CHK_CHARNUM		= "%1의 자리수는  %2 이어야합니다"
Const TYPE_POSITIVE			= "%1은 양수이어야합니다."


Function UNIGetMesg(Byval pTemp, Byval pMsg1, Byval pMsg2)
	If pMsg1 = "" Then
		UNIGetMesg = pTemp
	ElseIf pMsg1 <> "" And pMsg2 <> "" Then
		UNIGetMesg = Replace(Replace(pTemp, "%1", pMsg1), "%2", pMsg2)
	ElseIf pMsg1 <> "" And pMsg2 = "" Then
		UNIGetMesg = Replace(pTemp, "%1", pMsg1)
	End If
End Function

' -------- 문자열 공통 유틸리티 -----------
Function LeftH(ByVal strString, ByVal lngLength)
    Dim i, iLen, sTmp, iLen2, sRet
        
    For i = 1 To lngLength
        sTmp = Mid(strString, i, 1)
        sRet = sRet & sTmp
        
        If Asc(sTmp) < 0 Then    ' 한글 
            iLen2 = iLen2 + 2
        Else
            iLen2 = iLen2 + 1
        End If
        If iLen2 = lngLength Then Exit For
    Next
    'LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    LeftH = sRet
End Function
' -------- 문자/숫자 변환 -----------
' -- CHAR문자로 리턴 
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
	   pTmp = Replace(pTmp, "-", "")  '음수일경우 맨 앞자리에 표시 
	   pRet= "-" & Right(pRet & pTmp, (pLen-pDec)-1) & Right(pRet2 & pTmp2, pDec)
	End if  

	UNINumeric = pRet
End Function


' -- NUMERIC 리턴 
Function UNINumeric_old(Byval pData, Byval pLen, Byval pDec)
	Dim pTmp, pRet , pTmp2,  pRet2 , pMesg
	
	Err.Clear 
	On Error Resume Next
	
	   pTmp  = Replace(fix(unicdbl(pData,0)), ".", "")   ' 정수부분 
	   pTmp2 = Replace(unicdbl(pData,0) - pTmp, ".", "")  '소수부분 
	   
   
	
	If Err Then
		Call SaveHTFError(lgsPGM_ID, pData, UNIGetMesg(TYPE_CHK_NUM, pData, ""))
		Exit Function
	End If
	
	  pRet =  String(pLen, "0")
	  pRet2 = String(pDec, "0")
	 If unicdbl(pData,0) >= 0  then
	  
	   pRet = Right(pRet & pTmp, pLen - pDec) & Right(pTmp & pRet2, pDec)
	Else
	   pTmp = Replace(pTmp, "-", "")  '음수일경우 맨 앞자리에 표시 
	   pRet= "-" & Right(pRet & pTmp, (pLen-pDec)-1) & Right(pTmp & pRet2, pDec)
	End if  
	
	UNINumeric = pRet
End Function

Const TYPE_NOT_NULL = 0
Const TYPE_NULL		= 1
Const TYPE_NOT_NULL_DEFAULT_0 = 2

' -------- 널값 체크(문자,숫자,날짜) -----------
Function ChkNotNull(Byval pData, Byval pMesg)
	ChkNotNull = False
	If IsNull(pData) Or Trim(pData) = "" Then
		Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg(TYPE_CHK_NULL, pMesg,""))
		Exit Function
	End If
	ChkNotNull = True
End Function

' -- 날짜 형식 체크 
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

' -- 숫자검사 
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

' -- 8자리 날짜 리턴 
Function UNI8Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI8Date = String(8, " ")
		Exit Function
	End If
	UNI8Date = Year(pDate) & Right("0" & Month(pDate), 2) & Right("0" & Day(pDate), 2)
End Function

' -- 6자리 날짜 리턴 
Function UNI6Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI6Date = String(6, " ")
		Exit Function
	End If
	UNI6Date = Year(pDate) & Right("0" & Month(pDate), 2) 
End Function

' -- 4자리 날짜 리턴 
Function UNI4Date(Byval pDate)
	If IsNUll(pDate) Or pDate = "" Then 
		UNI4Date = String(4, " ")
		Exit Function
	End If
	UNI4Date = Year(pDate)
End Function



' -- 시작/종료 날짜 비교 
Function ChkDateFrTo(Byval pFrDt, Byval pToDt, Byval pMesg1, Byval pMesg2, Byval blnOver)
	ChkDateFrTo = False
	If pFrDt > pToDt Then	' -- 시작일이 종료일보다 크면 에러 
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_DATE_FRTO, pMesg1, pMesg2))
		Exit Function
	ElseIf DateDiff("m", pFrDt, pToDt) > 12 And blnOver = False Then	' -- 12개월보다 크면 
		Call SaveHTFError(lgsPGM_ID, pNum, UNIGetMesg(TYPE_CHK_DATE_OVER, "", ""))
		Exit Function
	End If	
	ChkDateFrTo = True
End Function

' -- Dash 제거(주민등록번호등)
Function UNIRemoveDash(Byval pData)
	UNIRemoveDash = Replace(pData, "-", "")
End Function

' -- 문자중 적당하지 않는거 검색 
Function ChkContents(ByVal pData, Byval pMesg)
    Dim i, iLen, sTmp, blnError
	On Error Resume Next
    blnError = False
    iLen = Len(pData)
    For i = 1 To iLen
        sTmp = Asc(Mid(pData, i, 1))
        If sTmp >= 65 And sTmp <= 90 Then    ' 한글은 음수/영어는 양수로 리턴 
            blnError = True
            Exit For   ' -- 영어포함스트링임 
        ElseIf sTmp >= 97 And sTmp <= 122 Then
            blnError = True
            Exit For   ' -- 영어포함스트링임 
        ElseIf sTmp = 32 Then
            If i < iLen And Asc(Mid(pData, i + 1, 1)) = 32 Then ' -- 공백은 1개일때만 허용 
                blnError = True
				Exit For   ' -- 공백 2개 
            End If
        ElseIf sTmp > 0 And sTmp < 32 Then
            blnError = True
            Exit For   ' -- 특수문자 
        ElseIf sTmp > 32 And sTmp < 45 Then     ' - 대쉬(45)는 허용 
            blnError = True
            Exit For   ' -- 특수문자 
        ElseIf sTmp > 45 And sTmp < 48 Then
            blnError = True
            Exit For   ' -- 특수문자 
        ElseIf sTmp > 57 And sTmp < 65 Then
            blnError = True
            Exit For   ' -- 특수문자 
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

' -- 일련번호 리턴: 000001, 000002
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

' -- 바운더리 체크 
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

' -- 법인 사업자등록번호에서 4:2 데이타 
Function GetRgstNo42(Byval pRgstNo)
	GetRgstNo42 = Mid(Replace(pRgstNo, "-", ""), 4, 2)
End Function


' -- 자리수 체크(데이터, 자리수)
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


' ---------------------- 홈텍스 파일 저장함수 --------------
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
		If lgblnPushDoc Then	' -- 저장된 파일 존재시 
		
			' 출력문서(pDesc)와 저장된문서(lgarrTAX_DOC)와 서식코드를 비교해 저장된 문서의 서식코드가 낮으면 출력한다.
			Dim pDoc1, pDoc2, i, iMaxCnt
			
			pDoc1 = Mid(pDesc, 4, 3)	' 4번째부터 3자리읽기 (예: 83A101....  101을 읽음)
			iMaxCnt = UBound(lgarrTAX_DOC)
			
			For i = 0 To iMaxCnt
				If Trim(lgarrTAX_DOC(i)) <> "" Then
				
					pDoc2 = Mid(lgarrTAX_DOC(i), 4, 3)	' 4번째부터 3자리읽기 
					If CDbl(pDoc1) > CDbl(pDoc2) Then
						If Err Then
							PrintLog "WriteLine2File lgarrTAX_DOC Error.. : " & Err.Description
							PrintLog "pDoc1=" & pDoc1
							PrintLog "pDoc2=" & pDoc2
							
							Exit Function
						End If
						' -- 출력해야될 서식의 코드가 기억된 서식의 코드보다 크다면,  기억된 서식 부터 출력한다.
						'PrintLog "lgarrTAX_DOC(" & i & ")=" & lgarrTAX_DOC(i)
						lgStream.WriteLine lgarrTAX_DOC(i)
						lgarrTAX_DOC(i) = ""
					End If
				End If
			Next
			
		End If
		
		lgStream.WriteLine pDesc		' -- 요청한 문서를 출력한다.
	End If
End Function

Function Write2File(Byval pDesc)
	On Error Resume Next
	If Not lgStream is Nothing Then
		If lgblnPushDoc Then	' -- 저장된 파일 존재시 
		
			' 출력문서(pDesc)와 저장된문서(lgarrTAX_DOC)와 서식코드를 비교해 저장된 문서의 서식코드가 낮으면 출력한다.
			Dim pDoc1, pDoc2, i, iMaxCnt
			
			pDoc1 = Mid(pDesc, 4, 3)	' 4번째부터 3자리읽기 (예: 83A101....  101을 읽음)
			iMaxCnt = UBound(lgarrTAX_DOC)
			
			For i = 0 To iMaxCnt
				If Trim(lgarrTAX_DOC(i)) <> "" Then
				
					pDoc2 = Mid(lgarrTAX_DOC(i), 4, 3)	' 4번째부터 3자리읽기 
					If uniCDbl(pDoc1,0) > UniCDbl(pDoc2,0) Then
						If Err Then
							PrintLog "Write2File lgarrTAX_DOC Error.. : " & Err.Description
							PrintLog "pDoc1=" & pDoc1
							PrintLog "pDoc2=" & pDoc2
							
							Exit Function
						End If
						' -- 출력해야될 서식의 코드가 기억된 서식의 코드보다 크다면,  기억된 서식 부터 출력한다.
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
		
		lgStream.Write pDesc		' -- 요청한 문서를 출력한다.
	End If
End Function

Function CloseFileSystem()

	' -- 기억된 서식중 남아있는게 있다면 출력후 종료한다 
	Dim iMaxCnt, i
	
	If lgblnPushDoc Then	' -- 저장된 파일 존재시 

		iMaxCnt = UBound(lgarrTAX_DOC)
				
		For i = 0 To iMaxCnt
			If Trim(lgarrTAX_DOC(i)) <> "" Then
				PrintLog "기억된 서식이 존재해 출력합니다 : " & lgarrTAX_DOC(i)	
				' -- 출력해야될 서식의 코드가 기억된 서식의 코드보다 크다면,  기억된 서식 부터 출력한다.  :** 기억된 서식의 순서가 다를경우는 다음버전에서 고민 ㅠㅠ 
				'PrintLog "lgarrTAX_DOC(" & i & ")=" & lgarrTAX_DOC(i)
							
				Select Case Mid(lgarrTAX_DOC(i), 3, 4)
					'Case "A172"	' -- 서식에서 엔터키를 안줄때..
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

Dim lgarrTAX_DOC(10)	' -- 파일순서대로 저장하지 않는 서식들 
Dim lgblnPushDoc
lgblnPushDoc =False		' -- 저장된 파일없음 

Function PushRememberDoc(Byval pDoc)
	Dim Index
	Index = GetPushIndex()
	If Index = -1 Then
		Err.Raise 60000, "inc_HomeTaxFunc.asp_PushRememberDoc()", "lgarrTAX_DOC 배열 초과"
	End If
	lgarrTAX_DOC(Index) = pDoc
	PrintLog "PushRememberDoc(" & Index & ")=" & pDoc
	lgblnPushDoc		= True	' -- 저장된 파일 있음 
End Function

Function GetPushIndex()	' -- 현재 비어 있는 배열인덱스 리턴 
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