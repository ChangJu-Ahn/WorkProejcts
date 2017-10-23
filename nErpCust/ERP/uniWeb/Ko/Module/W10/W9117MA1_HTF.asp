<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 국조제1호정상가격산출명세 
'*  3. Program ID           : W9117MA1
'*  4. Program Name         : W9117MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_KJ1

Set lgcTB_KJ1 = Nothing	' -- 초기화 

Class C_TB_KJ1
	' -- 테이블의 컬럼변수 
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
			 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		LoadData = False
			 
		PrintLog "LoadData IS RUNNING: "
			 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		lgStrSQL = ""
		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly)  = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If	
			LoadData = False  
			Exit Function  
		End If

		
		LoadData = True
	End Function

	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function

	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	End Function	
		
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "  A.W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10"
            lgStrSQL = lgStrSQL & " FROM TB_KJ1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	    AND A.W1 <> '4'"  & vbCrLf
			

			
			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

            lgStrSQL = lgStrSQL & " ORDER BY  A.W1 ASC" & vbcrlf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9117MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9117MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows, iSeqNo
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9117MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9117MA1"
	
	Set lgcTB_KJ1 = New C_TB_KJ1		' -- 해당서식 클래스 
	
	If Not lgcTB_KJ1.LoadData	Then Exit Function		' -- 제1호 서식 로드 
		
	'==========================================
	' -- 국조제1호정상가격산출명세 및 오류검증 
	
	iSeqNo = 1

	For iDx = 2 To 10
		
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' 일련번호 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx),stmp & "국외특수관계자_법인명(상호)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 60)
		stmp =lgcTB_KJ1.GetData("W" & iDx)
	
		lgcTB_KJ1.MoveNext
		lgcTB_KJ1.MoveNext  
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx),stmp & "국외특수관계자_소재국가") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 3)
		
		lgcTB_KJ1.MoveNext 
			
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "국외특수관계자_대표자(성명)") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 30)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "국외특수관계자_업종") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 7)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_KJ1.GetData("W" & iDx), stmp & "국외특수관계자_신고인과의관계") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 1)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp & "국외특수관계자_소재지") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 70)
		
		sHTFBody = sHTFBody & UNIChar("", 17) & vbCrLf	' -- 공란 
	
		lgcTB_KJ1.MoveFirst

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' 마지막이 아닐때, 값이 없으면 루프탈출 
			If lgcTB_KJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	
	Next
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	
	blnError = False : sHTFBody = "" : lgcTB_KJ1.MoveFirst 
	
	' -- 주식수변동상황 
	iSeqNo = 1	

	For iDx = 2 To 10
        Call lgcTB_KJ1.Find("W1='1'")	' 법인명 : 에러시 리턴하기 위해 : 2006.03.06 주석문추가 
        stmp =lgcTB_KJ1.GetData("W" & iDx)
        
		Call lgcTB_KJ1.Find("W1='9'")	' 대상거래부터 : 2006.03.06 : 화면에 코드/코드명으로 분리되면서 수정 

		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' 일련번호 
		
		If Not ChkBoundary("01,02,03,04,05,06,07,08,09,10", UNISeqNo(lgcTB_KJ1.GetData("W" & iDx),2), stmp  & "대상거래") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 2)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkBoundary("01,02,03,04,05,06", lgcTB_KJ1.GetData("W" & iDx), stmp  & "정상가격산출방법") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 2)
		
		lgcTB_KJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ1.GetData("W" & iDx), stmp  & "위의방법선택이유") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ1.GetData("W" & iDx), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 34) & vbCrLf	' -- 공란 
	
		lgcTB_KJ1.MoveFirst 

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' 마지막이 아닐때, 값이 없으면 루프탈출 
			If lgcTB_KJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	
	Next
	
		PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_KJ1 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9117MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
	
	End Select
	PrintLog "SubMakeSQLStatements_W9117MA1 : " & lgStrSQL
End Sub

%>
