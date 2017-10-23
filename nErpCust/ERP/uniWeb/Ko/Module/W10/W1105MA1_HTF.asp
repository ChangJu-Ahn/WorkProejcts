<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제3호의2(1)(2)표준손익계산서 
'*  3. Program ID           : W1105MA1
'*  4. Program Name         : W1105MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_3_3

Set lgcTB_3_3 = Nothing ' -- 초기화 

Class C_TB_3_3
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	
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
		Call GetData()
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
		Call GetData()
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFirst()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData()
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
		Else
			W1			= ""
			W2			= ""
			W3			= ""
			W4			= ""
			W5			= 0
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
				lgStrSQL = lgStrSQL & " FROM TB_3_3	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W1105MA1
	Dim A117
	Dim A102
	Dim A100
	Dim A129
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W1105MA1()
    Dim iKey1, iKey2, iKey3,dblAmt1
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1105MA1"

	Set lgcTB_3_3 = New C_TB_3_3		' -- 해당서식 클래스 
	
	If Not lgcTB_3_3.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	Set cDataExists = new TYPE_DATA_EXIST_W1105MA1
	Call lgcTB_3_3.Clone(oRs2)
	'==========================================
	' -- 제3호의2(1)(2)표준손익계산서 전자신고 및 오류검증 
	sHTFBody = "83"

	If lgcTB_3_3.W1 = "1" Then
		sHTFBody = sHTFBody & UNIChar("A115", 4) ' -- 일반법인용 
	Else
		sHTFBody = sHTFBody & UNIChar("A116", 4) ' -- 금융법인용 
	End If

	lgcTB_3_3.Find "W4 = '01'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '02'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '03'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '04'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '05'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '06'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '07'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '08'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '09'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '10'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '11'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '12'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '13'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '91'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '14'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '15'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '16'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '17'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '18'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '19'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '20'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '21'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '22'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '23'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '24'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '25'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '26'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '27'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '28'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '29'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '30'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '31'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '32'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '33'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '34'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '35'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '36'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '37'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '38'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '39'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '40'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '41'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '42'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '43'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '44'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '45'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '46'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '47'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '48'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '49'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '50'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '51'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '52'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '53'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '54'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '55'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '56'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '57'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '58'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '59'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '60'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '61'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '62'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '63'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '64'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '65'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '66'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '67'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '68'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '69'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '70'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '71'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '72'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '73'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '74'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '75'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '76'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '77'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '78'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '79'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '80'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '81'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '82'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '83'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '84'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '85'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '201'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '202'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '203'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '204'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '86'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '211'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '212'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '213'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '214'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '87'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '221'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '222'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '223'"	
	sHTFBody = sHTFBody & UNIChar(lgcTB_3_3.W3, 50)	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	lgcTB_3_3.Find "W4 = '224'"	
	If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)

	sHTFBody = sHTFBody & UNIChar("", 44)	' -- 공란 
	
	' -- 검증로직과 파일생성로직을 분리한다. (쿼리한 순서랑 파일 생성 순서가 틀리다)
	' -------------------------------------------------------------------------------	
	lgcTB_3_3.Find "W4 = '01'"	

	If lgcTB_3_3.W1 = "1" Then
		'sHTFBody = sHTFBody & UNIChar("A115", 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

		Do Until lgcTB_3_3.EOF 
			
		  SELECt Case  lgcTB_3_3.W4 
		     Case "01"
		            oRs2.MoveFirst
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '03'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
					
					'- 코드(01)매출액= 코드 02 + 03 + 04 + 05 + 06 + 07 + 08
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(01)매출액","코드 02 + 03 + 04 + 05 + 06 + 07 + 08"))
						   blnError = True	
					End If
			  Case "09"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '10'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
					
					'-  코드(09) 매출원가= 코드 10 + 14
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(09)매출원가","코드 10 + 14"))
						   blnError = True	
					End If	
			  
			  Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '91'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'- 코드(10)상품매출원가= 코드 11 + 12 - 13 - 91
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(10)상품매출원가","코드 11 + 12 - 13 - 91"))
						   blnError = True	
					End If	
			
			  Case "14"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '15'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '16'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  코드(14)제조,공사,임대,분양,운송,기타원가= 코드 15 + 16 - 17 - 18
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(14)제조,공사,임대,분양,운송,기타원가","코드 15 + 16 - 17 - 18"))
						   blnError = True	
					End If			
					
				Case "16"
			      
					
					'- 코드(16)당기총원가일반법인인 경우 부속명세서의 당기원가의 합과 일치16 = 
					'제조원가명세서(A117)의 
					'코드(34)(당기제품제조원가)           + 공사원가명세서(A118)의 코드(32)(당기공사원가)         
					'  + 임대원가명세서(A119)의 코드(17)(임대원가)           + 분양원가명세서(A120)의 코드(34)(당기완성주택등공사비)       
					'    + 운송원가명세서(A121)의 코드(30)(당기총운송원가)           + 기타원가명세서(A123)의 코드(32)(당기총원가 
					
					
			        	Set cDataExists.A117 = new C_TB_3_3_3	' -- W1107MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A117.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
									
							If Not cDataExists.A117.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "부속명세서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
						
							Else
								  cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '3' AND  W4 = '34' "	 ' 당기제품제조원가 
							     if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if
							      cDataExists.A117.Filter  ""
							      cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '4' AND  W4 = '32' "	 ' 당기공사원가 
							      if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
							      cDataExists.A117.Filter  ""
							      cDataExists.A117.MoveFist
							      cDataExists.A117.Filter "W1 = '5' AND  W4 = '17' "	 ' 임대원가 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								  cDataExists.A117.Filter "W1 = '6' AND  W4 = '34' "	 ' 당기완성주택등공사비 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								  cDataExists.A117.Filter " W1 = '7' AND  W4 = '30' "	 ' 당기총운송원가 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
								  cDataExists.A117.Filter  ""
								  cDataExists.A117.MoveFist
								   cDataExists.A117.Filter "W1 = '8' AND  W4 = '32' "	 ' 당기총운송원가 
								  if Not cDataExists.A117.Eof Then
									 sTmp = sTmp & UNICDbl(cDataExists.A117.W5,0)
								  End if	 
					
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A117 = Nothing				
	
							If UNICDbl(lgcTB_3_3.W5, 0) <> UNICDbl(sTmp,0) Then
								   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(16)당기총원가" ,"제조원가명세서(A117)의코드(34)(당기제품제조원가) " &_
								         "  + 공사원가명세서(A118)의 코드(32)(당기공사원가)  + 임대원가명세서(A119)의 코드(17)(임대원가)   " &_
								         "  + 분양원가명세서(A120)의 코드(34)(당기완성주택등공사비) " &_
								         "  + 운송원가명세서(A121)의 코드(30)(당기총운송원가)           + 기타원가명세서(A123)의 코드(32)(당기총원가"))
								   blnError = True	
							End If				
					
					
					
			    Case "19"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
				
					
					'-  코드(19)매출총이익= 코드 01 - 09	
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(19)매출총이익","코드 01 - 09"))
						   blnError = True	
					End If		
					
				Case "20"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '21'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '29'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '30'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '31'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '32'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '33'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '34'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '35'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					
				
					
					'- 코드(20)판매비와관리비= 코드 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29 + 30 + 31 + 32 + 33+ 34 + 35
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(20)판매비와관리비","코드21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29 + 30 + 31 + 32 + 33+ 34 + 35"))
						   blnError = True	
					End If	

				' -- 200603	: 서식 추가 
				Case "35"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '201'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '202'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '203'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '204'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- 코드(35)기타판매비와관리비= 코드 201+202+203+204
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(35) 기타판매비와관리비","코드 201 + 202 + 203 + 204"))
						   blnError = True	
					End If	
					
			   Case "36"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '19'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
				
					
					'- - 코드(36)영업이익= 코드 19 - 20			
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(36)영업이익","코드 19 - 20	"))
						   blnError = True	
					End If	
					
			  
					
					
						
			    Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '41'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '42'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '43'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '45'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '51'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- 코드(37)영업외수익= 코드 38 + 39 + 40 + 41 + 42 + 43 + 44 + 45 + 46 + 47 + 48 + 49 + 50+ 51 + 52			
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(37)영업외수익","코드 38 + 39 + 40 + 41 + 42 + 43 + 44 + 45 + 46 + 47 + 48 + 49 + 50+ 51 + 52		"))
						   blnError = True	
					End If		
					

				' -- 200603	: 서식 추가 
				Case "52"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '211'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '212'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '213'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '214'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- 코드(52) 기타영업외수익= 코드 211+212+213+214
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(52) 기타영업외수익","코드 211 + 212 + 213 + 214"))
						   blnError = True	
					End If	
								
		      Case "53"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '54'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '68'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '70'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- 코드(53)영업외비용= 코드 54 + 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62 + 63 + 64 + 65 + 66+ 67 + 68 + 69 + 70	
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(53)영업외비용","코드 54 + 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62 + 63 + 64 + 65 + 66+ 67 + 68 + 69 + 70"))
						   blnError = True	
					End If	

				' -- 200603	: 서식 추가 
				Case "70"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '221'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '222'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '223'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '224'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)

					'- 코드(52) 기타영업외비용= 코드 211+212+213+214
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(70) 기타영업외비용","코드 221 + 222 + 223 + 224"))
						   blnError = True	
					End If	
					
				Case "71"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '36'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '37'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  코드(71)경상이익= 코드 36 + 37 - 53
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(71)경상이익","코드 36 + 37 - 53"))
						   blnError = True	
					End If		
					
				Case "72"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '73'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '74'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '75'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '76'"	
					dblAmt1 = dblAmt1 - UNICDbl(oRs2("W5"), 0)
					
					
					'-  코드(72)특별이익= 코드 73 + 74 + 75 + 76
					
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(72)특별이익","코드 73 + 74 + 75 + 76"))
						   blnError = True	
					End If			
			  Case "77"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '78'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '79'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
				
					
					'- 코드(77)특별손실= 코드 78 + 79
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(77)특별손실","코드 19 - 20	"))
						   blnError = True	
					End If			
				Case "80"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '71'"	
					dblAmt1 = UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '72'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
					oRs2.Find "W4 = '77'"	
					dblAmt1 = dblAmt1- UNICDbl(oRs2("W5"), 0)
				
					
					'- 코드(80)법인세비용차감전순손익= 코드 71 + 72 - 77
					If UNICDbl(lgcTB_3_3.W5, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(80)법인세비용차감전순손익","코드 71 + 72 - 77	"))
						   blnError = True	
					End If		
					
				Case "81"
			        	Set cDataExists.A102 = new C_TB_15	' -- W1107MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A102.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
									
							' -- 제15호 과목별소득금액조정명세서(A102)서식 
							Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							Call SubMakeSQLStatements_W1105MA1("A102",iKey1, iKey2, iKey3)   
								
							cDataExists.A102.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A102.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A102.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "제15호 과목별소득금액조정명세서_손금산입및익금불산입", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
								'코드(81)법인세비용 : 소득금액조정합계표(A102)의 익금산입및손금불산입의 합계금액 보다 크면 오류(당기순이익과세법인이 아닌경우만 검증)
								If UNICDbl(lgcTB_3_3.W5, 0) > UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg(TYPE_CHK_HIGH_AMT, "코드(81)법인세비용","소득금액조정합계표(A102)의 익금산입및손금불산입의 합계금액"))
								End If
							End If
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A102 = Nothing
				Case "65"
				   if UNICDbl(lgcTB_3_3.W5, 0) >= 5000000 Then
				       	'-코드(65)기부금 > 500만원 인 경우기부금명세서(A129) 합계(999999) 금액이 <=0 오류(단,종류별구분이 비영리법인(50,60,70)인경우 검증제외)
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
						
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "제22호 기부금명세서(A129)", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
							
								  Call cDataExists.A129.Find ("2","W9_CD = '99'")
								If UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"), 0)<= 0Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_3.W5, UNIGetMesg("코드(65)기부금 > 500만원 인 경우기부금명세서(A129) 합계(999999) 금액이 <=0 으면 안됩니다", "",""))
								End If
							End If
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing
					End if						
					
					
				' -- 2006.03.17 개정: 라.기타에 금액이 있을경우, 가.나.다 금액이 0이면 오류.
				Case "204", "214", "224"
					If UNICDbl(lgcTB_3_3.W5, 0) <> 0 Then

		 				oRs2.MoveFirst
						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 3) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "코드(" & lgcTB_3_3.W4 & ") 라. 기타에 금액이 0이 아닐경우, 코드(" & oRs2("W4") & ") " & oRs2("W3") & " 금액이 0이면 오류입니다.")
						End If

						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 2) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "코드(" & lgcTB_3_3.W4 & ") 라. 기타에 금액이 0이 아닐경우, 코드(" & oRs2("W4") & ") " & oRs2("W3") & " 금액이 0이면 오류입니다.")
						End If

						oRs2.Find "W4 = '" & (UNICDbl(lgcTB_3_3.W4, 0) - 1) & "'"	
						If UNICDbl(oRs2("W5"), 0) <= 0 Then
							blnError = True
							Call SaveHTFError(lgsPGM_ID, oRs2("W5"), "코드(" & lgcTB_3_3.W4 & ") 라. 기타에 금액이 0이 아닐경우, 코드(" & oRs2("W4") & ") " & oRs2("W3") & " 금액이 0이면 오류입니다.")
						End If
					End If
			End Select 			
			
			
			
			'If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)
		
			lgcTB_3_3.MoveNext 
	   Loop
	
	Else
	
		' -- 금융/보험/증권업 법인용 : uniERP를 안쓰는 유형이라 검증로직이 없는듯보임 
	
		'sHTFBody = sHTFBody & UNIChar("A116", 4)	
		
		'Do Until lgcTB_3_3.EOF 
	
			'If Not ChkNotNull(lgcTB_3_3.W5, lgcTB_3_3.W3) Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3.W5, 16, 0)
		
			'lgcTB_3_3.MoveNext 
	   'Loop
	End If
	
	
	
	
	' ----------- 
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_3_3 = Nothing	' -- 메모리해제 
	
End Function


' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W1105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	   Case "A102" '-- 외부 참조 SQL
		
			' -- 소득금액조정합계표(A102)의 익금산입및손금불산입의 금액(2)의 합계와 일치 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '1'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf
	End Select
	PrintLog "SubMakeSQLStatements_W1105MA1 : " & lgStrSQL
End Sub
%>
