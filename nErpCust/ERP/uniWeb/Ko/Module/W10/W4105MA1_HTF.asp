
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제5호 특별비용조정명세서 
'*  3. Program ID           : W4105MA1
'*  4. Program Name         : W4105MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_5

Set lgcTB_5 = Nothing ' -- 초기화 

Class C_TB_5
	' -- 테이블의 컬럼변수 
	Dim SEQ_NO
	Dim W1_CD
	Dim W1
	Dim W2
	Dim W2_CD
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W7
	Dim W8
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 
		
		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		' 멀티행이지만 첫행을 리턴 
		Call GetData
		
		LoadData = True
	End Function

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.MoveFirst
		lgoRs1.Find pWhereSQL
		Call GetData
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
			SEQ_NO		= lgoRs1("SEQ_NO")
			W1_CD		= lgoRs1("W1_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_CD		= lgoRs1("W2_CD")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			W7			= lgoRs1("W7")
			W8			= lgoRs1("W8")
		Else
			SEQ_NO		= ""
			W1_CD		= ""
			W1			= ""
			W2			= ""
			W2_CD		= ""
			W3			= 0
			W4			= 0
			W5			= 0
			W6			= 0
			W7			= 0
			W8			= 0
		End If
	End Function
	
	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
	End Function
	
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
	End Sub

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_5 A WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W4105MA1
	Dim W4101MA1
	Dim W4103MA1
	Dim A140
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W4105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim arrdblSumAmt(8), dbl19Amt, dbl17Amt, dbl18Amt, dbl40Amt, dbl46Amt,dblW19Amt,dblW40Amt
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W4105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W4105MA1"

	Set lgcTB_5 = New C_TB_5		' -- 해당서식 클래스 
	
	If Not lgcTB_5.LoadData Then Exit Function			
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W4105MA1
	
	'==========================================
	' -- 제5호 특별비용조정명세서 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	' -- 주석: 국세청파일생성방식이 개정서식내용만 저장하는게 아니라 구서식의 데이타필드도 포함한다.
	
	lgcTB_5.Find "W2_CD = '01'"	' -- 중소기업투자준비금(2006.03 개정서식에 없다)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- 회사계상액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- 한도초과액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- 차감액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- 최저한세적용손금부인액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- 손금불산입계 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- 손금산입계 

	lgcTB_5.Find "W2_CD = '42'"	' -- 주권상장중소기업 등의 사업손실준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '02'"	' -- 연구인력개발준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '03'"	' -- 투용자손실준비금(2006.03 개정서식에 없다)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- 회사계상액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- 한도초과액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- 차감액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- 최저한세적용손금부인액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- 손금불산입계 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- 손금산입계 

	lgcTB_5.Find "W2_CD = '08'"	' -- 사회간접자본투자준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '16'"	' -- 기업구조조정전문회사(2006.03 개정서식에 없다)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- 회사계상액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- 한도초과액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- 차감액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- 최저한세적용손금부인액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- 손금불산입계 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- 손금산입계 

	lgcTB_5.Find "W2_CD = '45'"	' -- 부동산투자회사의 투자손실준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '17'"	' -- 100%손금산입고유목적 사업준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '18'"	' -- 부동산투자회사의 투자손실준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '14'"	' -- 유통개선지원준비금(2006.03 개정서식에 없다)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)	' -- 회사계상액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)	' -- 한도초과액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)	' -- 차감액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)	' -- 최저한세적용손금부인액 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)	' -- 손금불산입계 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)	' -- 손금산입계 

	lgcTB_5.Find "W2_CD = '43'"	' -- 주권상장기업 등의 자사주 처분손실준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '47'"	' -- 문화사업준비금 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '19'"	' -- 준비금 계 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '40'"	' -- 특별감가상각비 계 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '46'"	' -- 특례자산감가상각비 계 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '41'"	' -- 준비금 및 특별감가상각비 계 
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)

	lgcTB_5.Find "W2_CD = '44'"	' -- 사용자정의 
	sHTFBody = sHTFBody & UNIChar(lgcTB_5.W1, 50)
	If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
	If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
	If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
	If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
		If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
	End IF	
	If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
	If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)


	lgcTB_5.Find "W2_CD = '17'"	' -- 100%손금산입고유목적 사업준비금 
	If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)

	lgcTB_5.Find "W2_CD = '18'"	' -- 80%손금산입고유목적 사업준비금 
	If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 14)	' -- 공란 2006.03.07 수정 
	

	' 파일 생성후 검증 로직 수행 (파일 생성 로직은 주석처리요망)
	' ------------------------------------------------------------------------------------
	lgcTB_5.MoveFirst
	
	Do Until lgcTB_5.EOF 
	
		If lgcTB_5.W2_CD <> "" Then
			
			Select Case UNICDbl(lgcTB_5.W2_CD , 0)
				Case 01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43
						
					' -- 공통항목 
					' (5)차감액 = (3)회사계상액 - (4)한도초과액 
					If UNICDbl(lgcTB_5.W5, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W4, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(5)차감액","(3)회사계상액 - (4)한도초과액"))
					End If
			
					' (7)손금불산입계 = (4)한도초과액 + (6)최저한세적용손금부인 
					If UNICDbl(lgcTB_5.W7, 0) <> (UNICDbl(lgcTB_5.W4, 0) + UNICDbl(lgcTB_5.W6, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(7)손금불산입계","(4)한도초과액 + (6)최저한세적용손금부인"))
					End If
					
					' (8) 손금산입계 = (3)회사계상액 - (7)손금불산입계 
					If UNICDbl(lgcTB_5.W8, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W7, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(8)손금산입계","(3)회사계상액 - (7)손금불산입계"))
					End If
					
					arrdblSumAmt(3) = arrdblSumAmt(3) + UNICDbl(lgcTB_5.W3, 0)
					arrdblSumAmt(4) = arrdblSumAmt(4) + UNICDbl(lgcTB_5.W4, 0)
					arrdblSumAmt(5) = arrdblSumAmt(5) + UNICDbl(lgcTB_5.W5, 0)
					If lgcTB_5.W2_CD = "17" Or lgcTB_5.W2_CD = "18"  Then
					Else
						arrdblSumAmt(6) = arrdblSumAmt(6) + UNICDbl(lgcTB_5.W6, 0)
					End If
					arrdblSumAmt(7) = arrdblSumAmt(7) + UNICDbl(lgcTB_5.W7, 0)
					arrdblSumAmt(8) = arrdblSumAmt(8) + UNICDbl(lgcTB_5.W8, 0)
				Case 44
				   
				   
				   
				   'sHTFBody = sHTFBody & UNIChar(lgcTB_5.W1, 50)
				   'sHTFBody = sHTFBody & UNIChar(lgcTB_5.W2, 40)
				
				   ' -- 공통항목 
					' (5)차감액 = (3)회사계상액 - (4)한도초과액 
					If UNICDbl(lgcTB_5.W5, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W4, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(5)차감액","(3)회사계상액 - (4)한도초과액"))
					End If
			
					' (7)손금불산입계 = (4)한도초과액 + (6)최저한세적용손금부인 
					If UNICDbl(lgcTB_5.W7, 0) <> (UNICDbl(lgcTB_5.W4, 0) + UNICDbl(lgcTB_5.W6, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(7)손금불산입계","(4)한도초과액 + (6)최저한세적용손금부인"))
					End If
					
					' (8) 손금산입계 = (3)회사계상액 - (7)손금불산입계 
					If UNICDbl(lgcTB_5.W8, 0) <> (UNICDbl(lgcTB_5.W3, 0) - UNICDbl(lgcTB_5.W7, 0)) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(8)손금산입계","(3)회사계상액 - (7)손금불산입계"))
					End If
					
					arrdblSumAmt(3) = arrdblSumAmt(3) + UNICDbl(lgcTB_5.W3, 0)
					arrdblSumAmt(4) = arrdblSumAmt(4) + UNICDbl(lgcTB_5.W4, 0)
					arrdblSumAmt(5) = arrdblSumAmt(5) + UNICDbl(lgcTB_5.W5, 0)
					If lgcTB_5.W2_CD = "17" Or lgcTB_5.W2_CD = "18"  Then
					Else
						arrdblSumAmt(6) = arrdblSumAmt(6) + UNICDbl(lgcTB_5.W6, 0)
					End If
					arrdblSumAmt(7) = arrdblSumAmt(7) + UNICDbl(lgcTB_5.W7, 0)
					arrdblSumAmt(8) = arrdblSumAmt(8) + UNICDbl(lgcTB_5.W8, 0)
					
					
					
				Case 19
					' -- 준비금계 
					If UNICDbl(lgcTB_5.W3, 0) <> arrdblSumAmt(3) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(3)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
					
					If UNICDbl(lgcTB_5.W4, 0) <> arrdblSumAmt(4) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(4)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
					
					If UNICDbl(lgcTB_5.W5, 0) <> arrdblSumAmt(5) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(5)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
					
					If UNICDbl(lgcTB_5.W6, 0) <> arrdblSumAmt(6) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(6)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
					
					If UNICDbl(lgcTB_5.W7, 0) <> arrdblSumAmt(7) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(7)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
					
					If UNICDbl(lgcTB_5.W8, 0) <> arrdblSumAmt(8) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, lgcTB_5.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "준비금 계_(8)회사계상액","구분(01, 42, 02, 03, 08, 16, 45, 17, 18, 14, 43, 44)_(3)회사계상액의 합계"))
					End If
	
					' -- 200603 개정:  특별비용조정명세서(A108)서식의 코드(19)준비금계의 항목(5)최저한세적용손금부인액이최저한세적용계산서(A140)의 코드(5)준비금의 항목(4)조정감과 일치하는지 검증 추가 
					Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp 에 정의됨 
											
					' -- 추가 조회조건을 읽어온다.
					cDataExists.A140.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
					cDataExists.A140.WHERE_SQL = " AND A.W1 = '05' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
					If Not cDataExists.A140.LoadData() Then
	
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 최저한세적용계산서(A140) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
					Else

						If UNICDbl(lgcTB_5.W6, 0)  <> UNICDbl(cDataExists.A140.GetData("W4") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W6 & " <> " & cDataExists.A140.GetData("W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "제5호 특별비용조정명세서(A108)서식의 코드(19)준비금계의 항목(5)최저한세적용손금부인액","제4호 최저한세적용계산서(A140)의 코드(5)준비금의 항목(4)조정감"))
						End If
													
					End If

					' -- 사용한 클래스 메모리 해제 
					Set cDataExists.A140 = Nothing		
    
			End Select
			
			Select Case UNICDbl(lgcTB_5.W2_CD , 0)
				Case 01

					Set cDataExists.W4101MA1 = new C_TB_31_1	' -- W4101MA1_HTF.asp 에 정의됨 
								
					' -- 추가 조회조건을 읽어온다.
					cDataExists.W4101MA1.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
					cDataExists.W4101MA1.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
					If Not cDataExists.W4101MA1.LoadData() Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 세액공제신청서(A165) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
					Else

						If UNICDbl(lgcTB_5.W3, 0) > 0 AND UNICDbl(lgcTB_5.W3, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W5") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "중소기업투자준비금_회사계상액","제31호(1)중소기업투자준비금조정_(5)회사계상액"))
						End If

						If UNICDbl(lgcTB_5.W4, 0) > 0 AND UNICDbl(lgcTB_5.W4, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W6") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "중소기업투자준비금_한도초과액","제31호(1)중소기업투자준비금조정_(6)한도초과액"))
						End If

						If UNICDbl(lgcTB_5.W6, 0) > 0 AND UNICDbl(lgcTB_5.W6, 0) <> UNICDbl(cDataExists.W4101MA1.GetData(1,"W7") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "중소기업투자준비금_최저한세적용손금부인액","제31호(1)중소기업투자준비금조정_(7)최저한세적용에따른손금부인액"))
						End If												
					End If

					' -- 사용한 클래스 메모리 해제 
					Set cDataExists.W4101MA1 = Nothing		
					
					

				Case 02

					Set cDataExists.W4103MA1 = new C_TB_31_2	' -- W4103MA1_HTF.asp 에 정의됨 
								
					' -- 추가 조회조건을 읽어온다.
					cDataExists.W4103MA1.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
					cDataExists.W4103MA1.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
					If Not cDataExists.W4103MA1.LoadData() Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 세액공제신청서(A165) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
					Else

						If UNICDbl(lgcTB_5.W3, 0) > 0 AND UNICDbl(lgcTB_5.W3, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W4") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "연구및인력개발준비금_회사계상액","제31호(2)연구및인력개발준비금_(4)회사계상액"))
						End If

						If UNICDbl(lgcTB_5.W4, 0) > 0 AND UNICDbl(lgcTB_5.W4, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W5") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "연구및인력개발준비금_한도초과액","제31호(2)연구및인력개발준비금_(5)한도초과액"))
						End If

						If UNICDbl(lgcTB_5.W6, 0) > 0 AND UNICDbl(lgcTB_5.W6, 0) <> UNICDbl(cDataExists.W4103MA1.GetData(1,"W6") , 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "연구및인력개발준비금_최저한세적용손금부인액","제31호(2)연구및인력개발준비금_(6)최저한세적용에따른손금부인액"))
						End If												
					End If

					' -- 사용한 클래스 메모리 해제 
					Set cDataExists.W4103MA1 = Nothing		
				
				Case 19
					dbl19Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 17
					dbl17Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 18
					dbl18Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 40
					dbl40Amt = 	UNICDbl(lgcTB_5.W5, 0)
				Case 46
					dbl46Amt = 	UNICDbl(lgcTB_5.W5, 0)
					
			End Select
			
			If Not ChkNotNull(lgcTB_5.W3, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W3, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W4, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W4, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W5, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W5, 15, 0)
		    
		    
		    If lgcTB_5.W2_CD <> "17" And lgcTB_5.W2_CD <> "18"  Then   '최저한세없음 
					
				If Not ChkNotNull(lgcTB_5.W6, lgcTB_5.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W6, 15, 0)
			End IF	
					
			If Not ChkNotNull(lgcTB_5.W7, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W7, 15, 0)
				
			If Not ChkNotNull(lgcTB_5.W8, lgcTB_5.W1) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_5.W8, 15, 0)
		End If
		
		lgcTB_5.MoveNext 
	Loop

	' -- 점검 
	If dblW19Amt - dbl17Amt - dbl18Amt > 0 Then
	
		Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp 에 정의됨 
								
		' -- 추가 조회조건을 읽어온다.
		cDataExists.A140.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A140.WHERE_SQL = " AND A.W1 = '05' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
		If Not cDataExists.A140.LoadData() Then
	
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 최저한세적용계산서(A140) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else

			If dblW19Amt - dbl17Amt - dbl18Amt <> UNICDbl(cDataExists.A140.GetData("W3") , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "연구및인력개발준비금_회사계상액","제31호(2)연구및인력개발준비금_(4)회사계상액"))
			End If
										
		End If

		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A140 = Nothing		
	End If

	If dblW40Amt + dbl46Amt > 0 Then
	
		Set cDataExists.A140 = new C_TB_4	' -- W6127MA1_HTF.asp 에 정의됨 
								
		' -- 추가 조회조건을 읽어온다.
		cDataExists.A140.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A140.WHERE_SQL = " AND A.W1 = '06' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
		If Not cDataExists.A140.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 최저한세적용계산서(A140) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else

			If dblW19Amt - dbl17Amt - dbl18Amt <> UNICDbl(cDataExists.A140.GetData("W3") , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_5.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "연구및인력개발준비금_회사계상액","제31호(2)연구및인력개발준비금_(4)회사계상액"))
			End If
										
		End If

		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A140 = Nothing		
	End If

	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_5 = Nothing	' -- 메모리해제 
	
End Function


%>
