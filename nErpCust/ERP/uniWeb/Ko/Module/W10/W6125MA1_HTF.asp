<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제8호(갑) 공제감면세액 명세서 
'*  3. Program ID           : W6125MA1
'*  4. Program Name         : W6125MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_8A

Set lgcTB_8A = Nothing ' -- 초기화 

Class C_TB_8A
	' -- 테이블의 컬럼변수 
	Dim W_TYPE
	Dim W1_CD
	Dim W1
	Dim W2
	Dim W2_1
	Dim W3
	Dim W4
	Dim W7
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim SELECT_SQL		' -- 리턴 필드 변경 
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
			W_TYPE		= lgoRs1("W_TYPE")
			W1_CD		= lgoRs1("W1_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_1		= lgoRs1("W2_1")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W7			= lgoRs1("W7")
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

				If SELECT_SQL = "" Then
					lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & SELECT_SQL & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_8A A WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				lgStrSQL = lgStrSQL & "	ORDER BY W_TYPE , W1_CD "
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6125MA1
	Dim A161
	Dim A174
	Dim A165
	Dim A101
	Dim A179
	Dim A181
	Dim A175
	Dim A149
	Dim A154
	Dim A151
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6125MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim dblAmt, dblSum1Amt, dblSum2Amt, dblSum3Amt, dbl10Amt, dbl66Amt, dbl30Amt, dbl49Amt, dbl70Amt, dbl80Amt, dbl50Amt
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6125MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6125MA1"

	Set lgcTB_8A = New C_TB_8A		' -- 해당서식 클래스 
	
	If Not lgcTB_8A.LoadData Then Exit Function			
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W6125MA1
	
	'==========================================
	' -- 제8호(갑) 공제감면세액 명세서 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	' -- 2006.03.07 수정: 화면순서와 파일생성 순서가 다른다 
	
	lgcTB_8A.Find "W2_1 = '01'"	' -- 외국납부세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '02'"	' -- 재해손실세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '71'"	' -- 농업소득세세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '03'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '90'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '08'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '09'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '04'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '07'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '94'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '86'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '87'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = 88''"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '72'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '73'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '81'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '82'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '61'"	' -- 세액공제 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "코드61 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)

	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '10'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- 그리드2 
	lgcTB_8A.Find "W2_1 = '11'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '74'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '12'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '13'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '23'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '14'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '16'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '17'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '18'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '19'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '24'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '97'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '98'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '64'"	' -- 세액공제 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "코드64 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)
	
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '30'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- 그리드 3
	lgcTB_8A.Find "W2_1 = '31'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '93'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '75'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '32'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '85'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '34'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '76'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '35'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '36'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '77'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '37'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '91'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '42'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '95'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '96'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '65'"	' -- 세액공제 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "코드65 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)
	
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '49'"	' -- 세액공제2_소계_공제세액 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '50'"	' -- 세액공제2_합계_공제세액 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '51'"	' -- 공제감면세액총계 
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '83'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '89'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	' -- 개정서식후 추가된것들(ㅠ.ㅠ)
	' -- 그리드1에 추가된것들 
	lgcTB_8A.Find "W2_1 = '57'"	' -- 경제자유구역 개발사업시행자 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '58'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '59'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '60'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '67'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '66'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '99'"	' -- 세액공제 
	If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
		If Not ChkNotNull(lgcTB_8A.W1, "코드99 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)

	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	' -- 그리드3에 추가된것 
	lgcTB_8A.Find "W2_1 = '84'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
	If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)

	lgcTB_8A.Find "W2_1 = '70'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	lgcTB_8A.Find "W2_1 = '80'"	' -- 세액공제 
	If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 14)	' -- 공란	2006.03.07 수정 
	

	' -- 파일 생성후 검증로직 수행한다	
	'lgcTB_8A.MoveFirst		' -- 이 로직으로는 이상한 곳으로 이동한다. 2006.03.09
	lgcTB_8A.Find "W2_1 = '03'"	' -- 세액공제 
	' -----------------------------------------------------------------
	
	Do Until lgcTB_8A.EOF 
	
		Select Case lgcTB_8A.W_TYPE 
			Case "0"	' W3, W4 출력 
				
				' -- 데이타 저장 
				If lgcTB_8A.W2_1 = "61" Then		' -- 사용자가 구분을 지정할수 있는 코드 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "코드61 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
					End If
					
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)
				End If	
				if lgcTB_8A.W2_1 <> 10 then '소계 
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
				End if
				If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
			
				' -- 각 코드별 점검 
				Select Case lgcTB_8A.W2_1
					Case "02"
						'공제감면세액계산서(1)(A161)의 재해손실세액공제 항목(4)재해손실세액공제의 공제감면세액과 일치 
						'(코드(2)의 항목(4)가 “0”보타 큰 경우 A161 반드시 입력)
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- 감면세액 
					
							Set cDataExists.A161 = new C_TB_8_1	' -- W6124MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A161.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A161.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A161.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_공제세액이 0 보다 큰 경우 공제감면세액계산서(1)(A161) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A161.GetData("W4_2") , 0) Then
									blnError = True
								
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A161 = Nothing		
					
						End If
						
					Case "90" ' -- 고용창출형창업기업에 대한 세액감면_대상세액(코드90)

						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- 대상세액 
					
							'코드(90)의 대상세액이 "0"보다 큰 경우 고용창출형창업기업감면세액계산서(A174)의 서식유무 검증 
							
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(90)의 고용창출형창업기업에 대한 세액감면 대상세액이 0보다 큰 경우 고용창출형창업기업감면세액계산서(A174)의 서식이 있어야합니다.해당서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If

					Case "08"
					   
					   '코드(08)의 대상세액이 0보다 큰 경우 공장및본사를 수도권생활지역외의 지역으로 이전하는 법인에 대한 임시특별세액감면신청서(A207)의 서식유무 검증 
					    If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- 대상세액 
						
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(08)의 대상세액이 0보다 큰 경우 공장및본사를 수도권생활지역외의 지역으로 이전하는 법인에 대한 임시특별세액감면신청서(A207)의 서식이 있어야합니다. 해당서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If
				
						
					Case "04" '
					    '코드(04)의 대상세액이 0보다 큰 경우 영농조합법인세액면제신청서(A208)의 서식유무 검증 
					
						
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then
					
						
						
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(04)의 대상세액이 0보다 큰 경우 영농조합법인세액면제신청서(A208)의  서식이 있어야합니다.해당서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If
					
					Case "09" ' -- 수도권생활지역외지역 이전본사에 대한 세액감면_대상세액(코드09)
					    '코드(09)의 대상세액이 0보다 큰 경우 공장및본사를 수도권생활지역외의 지역으로 이전하는 법인에 대한 임시특별세액감면신청서(A207)의 서식유무 검증 
						
						
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(09)의 대상세액이 0보다 큰 경우 공장및본사를 수도권생활지역외의 지역으로 이전하는 법인에 대한 임시특별세액감면신청서(A207)의  서식이 있어야합니다.해당 서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If
					

					Case "07" ' -- 영어조합법인 대한 세액감면_대상세액(코드07)
						'- 코드(07)의 대상세액이 0보다 큰 경우 영어조합법인세액면제신청서(A209)의 서식유무 검증 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- 대상세액 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(07)의 대상세액이 0보다 큰 경우 영어조합법인세액면제신청서(A209)의  서식이 있어야합니다.해당 서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If
					
			

					Case "94" ' -- 지급조서전자제출 대한 세액감면_대상세액(코드94)
					    ' - 코드(94)의 대상세액이 "0"보다 큰 경우 지급조서전자제출공제세액명세서(A222)의 서식유무 검증 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- 대상세액 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(94)의 대상세액이 0보다 큰 경우 지급조서전자제출공제세액명세서(A222)의  서식이 있어야합니다.해당 서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						End If	

					' -- 200603 개정서식 반영 
					Case "70"
						dbl70Amt = dblSum1Amt	
						
						If UNICDbl(lgcTB_8A.W4, 0) <> dbl70Amt Then	' -- 감면세액 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & dbl70Amt, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계(코드70)_감면세액","코드(03+90+08+09+04+07+86+87+88+57+58+59+60+67+72+73+81+82+61)_감면세액"))
						End If	
						lgcTB_8A.W4 = 0	: dblSum1Amt = 0	' -- 루프에서 썸값을 구하는걸 초기화 

					' -- 200603 개정서식 반영 
					Case "80"

						dbl80Amt = dblSum1Amt

						If UNICDbl(lgcTB_8A.W4, 0) <> dbl80Amt Then	' -- 감면세액 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & dbl80Amt & ";" & dblSum1Amt, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계(코드80)_감면세액","코드(01+02+71+66+94+99)_감면세액"))
						End If	
					
						lgcTB_8A.W4 = 0	: dblSum1Amt = 0
					' -- 200603 개정서식 반영 
					Case "10" ' -- 소계 
						If UNICDbl(lgcTB_8A.W4, 0) <> (dbl70Amt + dbl80Amt) Then	' -- 감면세액 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4 & "<>" & (dbl70Amt + dbl80Amt), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계(코드10)_감면세액","코드(70+80)_감면세액"))
						End If	
						
						dbl10Amt = 	UNICDbl(lgcTB_8A.W4, 0)		' -- 코드10+코드66 금액과 3호 비교는 젤 마지막에 루프밖에서 한다.	
					
					Case "66"
						dbl66Amt = 	UNICDbl(lgcTB_8A.W3, 0)																
				End Select
				
				dblSum1Amt = dblSum1Amt + UNICDbl(lgcTB_8A.W4, 0)	' 소계_감면세액 
				
			Case "1"
				' -- 데이타 저장 
				If lgcTB_8A.W2_1 = "64" Then		' -- 사용자가 구분을 지정할수 있는 코드 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "코드64 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
					End If
									
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)
				End If	
				If  lgcTB_8A.W2_1 <> "30" Then
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
				End If	
				
				If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
			
				' -- 각 코드별 점검 
				Select Case lgcTB_8A.W2_1
				
					Case "16"	' -- 지방이전중소기업감면_대상세액 
					   '- 코드(16)의 대상세액이 0보다 큰 경우 수도권과밀억제권역 외의 지역으로 이전하는 중소기업감면세액계산서(A206)의 서식유무 검증 
						If UNICDbl(lgcTB_8A.W3, 0) > 0 Then	' -- 대상세액 
					
							blnError = True

							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg("코드(16)의 대상세액이 0보다 큰 경우 수도권과밀억제권역 외의 지역으로 이전하는 중소기업감면세액계산서(A206)의  서식이 있어야합니다.해당 서식이 존재하지 않으면 UNIERP 팀에 문의하시기 바랍니다", "",""))
						
						End If	

					

					
					Case "30"	' -- 소계		
						If UNICDbl(lgcTB_8A.W4, 0) <> dblSum2Amt Then	' -- 감면세액 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계(코드30)_감면세액","코드(11+74+12+13+23+14+16+17+18+19+24+97+98+64)_감면세액"))
						End If	
						dbl30Amt = 	UNICDbl(lgcTB_8A.W4, 0)										
				End Select
				
				dblSum2Amt = dblSum2Amt + UNICDbl(lgcTB_8A.W4, 0)	' 소계_감면세액 
				
				
			Case "2"		' W3, W4, W7 출력 
				If lgcTB_8A.W2_1 = "65"  Then		' -- 사용자가 구분을 지정할수 있는 코드 
				
					If UNICDbl(lgcTB_8A.W3, 0) <> 0 Or UNICDbl(lgcTB_8A.W4, 0) <> 0 Or UNICDbl(lgcTB_8A.W7, 0) <> 0 Then
						If Not ChkNotNull(lgcTB_8A.W1, "코드" & lgcTB_8A.W2_1 & " 구분") Then blnError = True	' -- 금액이 0아니면 Not Null
					End If
								
					'sHTFBody = sHTFBody & UNIChar(lgcTB_8A.W1, 50)	' -- 사용자 정의 구분(Null 허용)
				End If	
				
				If lgcTB_8A.W2_1 ="49" Or lgcTB_8A.W2_1 = "50" Or lgcTB_8A.W2_1 = "51" Then	
				Else
					If Not ChkNotNull(lgcTB_8A.W3, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W3, 15, 0)
						
					If Not ChkNotNull(lgcTB_8A.W4, lgcTB_8A.W1) Then blnError = True	
					'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W4, 15, 0)
				End if
				
				If Not ChkNotNull(lgcTB_8A.W7, lgcTB_8A.W1) Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8A.W7, 15, 0)
				
				' -- 5 + 6 > 7
				If UNICDbl(lgcTB_8A.W3, 0) + UNICDbl(lgcTB_8A.W4, 0) < UNICDbl(lgcTB_8A.W7, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W3, UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT,  lgcTB_8A.W1 &  "공제세액","전기이월액+당기발생액"))
				End If
				
				' -- 각 코드별 점검 
				Select Case lgcTB_8A.W2_1
				
					
					Case "75" ' -- 기업의어음제도개선을위한 세액공제_당기발생 
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- 대상세액 
					
								

							Set cDataExists.A179 = new C_TB_JT2_2	' -- W6109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A179.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A179.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A179.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 기업의어음제도개선을위한공제세액계산서(A179) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A179.GetData("W13") , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기업의어음제도개선을위한 세액공제","기업의어음제도개선을위한공제세액계산서(A179)의 항목(13)공제세액"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A179 = Nothing							
						End If	
					
' -- 2006.03.27 검증삭제	
'					Case "32" ' -- 연구인력개발비세액공_당기발생 
'						PrintLog "dbl66Amt : " & dbl66Amt
'						' -- 200603 개정: 66코드가 위로 올라감 
'						If UNICDbl(lgcTB_8A.W4, 0) + dbl66Amt > 0 Then	' -- 대상세액 
'					      '- 코드(32)의 공제세액이"0"보다 큰 경우 연구및인력개발비발생명세서(A181)의 서식유무 검증 
'
'							Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp 에 정의됨 
'								
'							' -- 추가 조회조건을 읽어온다.
'							cDataExists.A181.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
'							cDataExists.A181.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
'								
'							If Not cDataExists.A181.LoadData() Then
'								blnError = True
'								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 연구및인력개발비발생명세서(A181) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
'						
'							End If
'					
'							
'							' -- 사용한 클래스 메모리 해제 
'							Set cDataExists.A181 = Nothing							
'						End If	
						
						
						
					Case "37" ' 
						If UNICDbl(lgcTB_8A.W4, 0) <> 0 Then	' -- 대상세액 
					
						 '- 코드(37)의 공제세액이 0 이 아닌경우 농어촌특별세 과세대상 감면세액합계표 (A151)의 항목(137)임시투자세액공제 항목(4)감면세액 과 일치 

							Set cDataExists.A151 = new C_TB_13	' -- W6111MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A151.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A151.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A151.LoadData() Then
								Response.End 
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 농어촌특별세 과세대상 감면세액합계표 (A151) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							ElSE
							   '감면세액합계표 (A151)의 항목(137)임시투자세액공제 항목(4)감면세액 과 일치 
							   cDataExists.A151.FIND 1, " W2_CD='137' "	
							   if UNICDbl(lgcTB_8A.W4, 0) <>  UNIcdbl(cDataExists.A151.GetData(1, "W4") ,0) Then
							      blnError = True
								  Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "임시투자세액공제액","감면세액합계표 (A151)의 항목(137)임시투자세액공제 항목(4)감면세액"))
							   End if
						
							End If
	
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A151 = Nothing							
						End If	
					Case "91" ' -- 고용증대특별세액공제_당기발생 
						If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- 대상세액 
					
						

							Set cDataExists.A175 = new C_TB_JT11_5	' -- W6113MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A175.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A175.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A175.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 고용증대특별세액공제 공제세액계산서(A175) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
								If UNICDbl(lgcTB_8A.W4, 0) <> UNICDbl(cDataExists.A175.GetData("W10") , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "고용증대특별세액공제","고용증대특별세액공제 공제세액계산서(A175)의 항목(15)공제세액"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A175 = Nothing							
						End If							
						
					' -- 200603 개정: 66코드가 위로 올라감..	
'					Case "66"
'						dbl66Amt = 	UNICDbl(lgcTB_8A.W7, 0)		' -- 코드10+코드66 금액과 3호 비교는 젤 마지막에 루프밖에서 한다.	
'					 If UNICDbl(lgcTB_8A.W4, 0) > 0 Then	' -- 대상세액 
'					
'						
'
'							Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp 에 정의됨 
'								
'							' -- 추가 조회조건을 읽어온다.
'							cDataExists.A181.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
'							cDataExists.A181.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
'								
'							If Not cDataExists.A181.LoadData() Then
'								blnError = True
'								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 기업의어음제도개선을위한공제세액계산서(A179) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
'						
'							End If
'					
'							
'							' -- 사용한 클래스 메모리 해제 
'							Set cDataExists.A179 = Nothing							
'						End If	
						
						
					
					Case "49"	' --소계 
					
					
						If UNICDbl(lgcTB_8A.W7, 0) <> UNICDbl(dblSum3Amt,0) Then	' -- 공세액 
							blnError = True
							Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_8A.W7, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계(코드30)_감면세액","코드(31+93+75+32+85+34+76+35+36+77+37+91+42+95+84+96+65)_감면세액"))
						End If		

						If UNICDbl(lgcTB_8A.W7, 0) > 0 Then	' -- 공제세액 
						
						
						   ' 세액공제조정명세서(3)(A149)의 항목(116)공제세액의 합계와 일치(코드(49)의 공제세액이 0 보다 큰 경우 A149 반드시 입력)
					
							Set cDataExists.A149 = new C_TB_8_3	' -- W6103MA1_HTF.asp 에 정의됨 
								
							
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A149.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
						'	cDataExists.A149.WHERE_SQL = " AND A.W_CODE = '30' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
						

							If Not cDataExists.A149.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_8A.W1 & "_대상세액이 '0'보다 큰 경우 세액공제조정명세서(3)(A149) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
								
								' -- 일치체크하지 않고 서식유무만 체크함: 2006.03.10 개정반영 
							'Else 
							     'cDataExists.A149.FIND 2, "SEQ_NO = 9999999 "
					
							     'if UNICDbl(lgcTB_8A.W7, 0)  <>   UNICDbl(cDataExists.A149.GetData(2, "C_W116"),0)  Then
							     '  	blnError = True
								'	Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제세액_소계","세액공제조정명세서(3)(A149)의 항목(116)공제세액의 합계"))
							     'End If
							   	
							End If
					
						End If	
					
						dbl49Amt = 	UNICDbl(lgcTB_8A.W7, 0)
						Set cDataExists.A149 = Nothing			
						
					' -- 200603 개정서식 적용 		
					Case "50"	 ' -- 합계		
						' -- 코드(30)공제세액_소계 + 코드(49)공제세액_소계 
						If dbl30Amt + dbl49Amt <> UNICDbl(lgcTB_8A.W7, 0) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(50)공제세액_소계","코드(30)공제세액_소계 + 코드(49)공제세액_소계"))
						End If

						'합계_공제세액(코드50) - 코드(66)공제대상세액은 법인세과세표준및세액조정계산서(A101)의 코드(17)공제감면세액(ㄱ)을 기입(일치) 
						Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp 에 정의됨 
												
						' -- 추가 조회조건을 읽어온다.
						'cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
						'cDataExists.A101.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
						If Not cDataExists.A101.LoadData() Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, "", "소계_감면세액(코드50) - 코드(66)공제대상세액이 '0'보다 큰 경우 제3호법인세과세표준및세액조정계산서(A101) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
						Else
							If UNICDbl(lgcTB_8A.W7, 0) <> UNICDbl(cDataExists.A101.W17, 0) Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID,  UNICDbl(lgcTB_8A.W7, 0) - dbl66Amt , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계_감면세액(코드50)","제3호법인세과세표준및세액조정계산서(A101)의 코드(17)공제감면세액(ㄱ)"))
							End If
						End If
														
						' -- 사용한 클래스 메모리 해제 
						Set cDataExists.A101 = Nothing	

						dbl50Amt = UNICDbl(lgcTB_8A.W7, 0)
					' -- 200603 개정서식 적용 		
					Case "51"	 ' -- 합계		

						If UNICDbl(lgcTB_8A.W7, 0) <> (dbl10Amt + dbl50Amt) Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID,  UNICDbl(lgcTB_8A.W7, 0) - dbl66Amt , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제감면세액총계(51)","코드(10+50)의 공제세액"))
						End If
						
					Case "83" ' -- 기술도입대가에 대한조세면제_공제세액 
						If UNICDbl(lgcTB_8A.W7, 0) > 0 Then	' --공제세 

							' 미개발 : 서식미개발(기술도입대가에대핝조세면제명세서(A154))
						End If																									
				End Select
				
				dblSum3Amt = dblSum3Amt + UNICDbl(lgcTB_8A.W7, 0)	' 소계_감면세액 
		End Select
		
		lgcTB_8A.MoveNext 
	Loop

	' -- 루프밖 검증부분 
	If dbl10Amt > 0 Then		' -- 200603 개정서식 적용 
	
		'소계_감면세액(코드10) + 코드(66)공제대상세액은 법인세과세표준및세액조정계산서(A101)의 코드(19)공제감면세액(ㄴ)을 기입(일치)
		Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp 에 정의됨 
								
		' -- 추가 조회조건을 읽어온다.
		'cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		'cDataExists.A101.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
		If Not cDataExists.A101.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "소계_감면세액(코드10) '0'보다 큰 경우 제3호법인세과세표준및세액조정계산서(A101) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
			If dbl10Amt <> UNICDbl(cDataExists.A101.W19 , 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_8A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소계_감면세액(코드10)","제3호법인세과세표준및세액조정계산서(A101)의 코드(19)공제감면세액(ㄴ)"))
			End If
		End If
										
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A101 = Nothing	
	End If
								
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_8A = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6125MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A166" '-- 외부 참조 SQL
			lgStrSQL = ""

	End Select
	PrintLog "SubMakeSQLStatements_W7109MA1 : " & lgStrSQL
End Sub
%>
