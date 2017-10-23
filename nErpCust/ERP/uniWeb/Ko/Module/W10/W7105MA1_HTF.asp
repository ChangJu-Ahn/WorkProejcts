<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제50호 자본금과 적립금조정명세서(갑)
'*  3. Program ID           : W7105MA1
'*  4. Program Name         : W7105MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_50A
dim lgcTB_3_2_2
dim lgcCOMPANY_HISTORY

Set lgcTB_50A = Nothing ' -- 초기화 
Set lgcTB_3_2_2 = Nothing ' -- 초기화 

'---------------------------------------------------------------------
'---------------------------------------------------------------------

Class C_TB_50A
	' -- 테이블의 컬럼변수 
	Dim W_CD
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W_DESC
	
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
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W_CD		= lgoRs1("W_CD")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W_DESC		= lgoRs1("W_DESC")
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
				lgStrSQL = lgStrSQL & " FROM TB_50A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND NOT( W_CD BETWEEN '171' AND '177' )" '200703 TEMP
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class



'---------------------------------------------------------------------
'---------------------------------------------------------------------





  
Class C_TB_3_2_2
	' -- 테이블의 컬럼변수 

	Dim W1 
	Dim W2
	Dim W3
	Dim W4

	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3,iKey4
				 
		On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		if lgcCOMPANY_HISTORY.COMP_TYPE2="1" then '일반 
			iKey4="41,44,50,57"
		else
			iKey4="68,69,73,79"
		end if

		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3,iKey4)                                       '☜ : Make sql statements

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
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
		
			W1			= lgoRs1("W1") '자본금 
			W2			= lgoRs1("W2") '자본잉여금 
			W3			= lgoRs1("W3") '이익잉여금 
			W4			= lgoRs1("W4") '자본조정 
			
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
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3,pCode4)
		dim tmpKey
		tmpKey = split(pCode4,",")
	    Select Case pMode 
	      Case "H"

				
				lgStrSQL = ""
	           
				lgStrSQL = lgStrSQL & " SELECT                       "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD =  " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(0)&"'                "
				lgStrSQL = lgStrSQL & "   AND W2 LIKE '3%' ) W1      "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR =" & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(1)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W2       "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " ( SELECT A.W6-A.W5 FROM           "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(2)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W3     "
				lgStrSQL = lgStrSQL & " ,                            "
				lgStrSQL = lgStrSQL & " (SELECT A.W6-A.W5 FROM            "
				lgStrSQL = lgStrSQL & " TB_3_2_2 A                   "
				lgStrSQL = lgStrSQL & "   WHERE a.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND a.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "   AND W4='"&tmpKey(3)&"'                "
				lgStrSQL = lgStrSQL & "  AND W2 LIKE '3%'  ) W4      "

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class







' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W7105MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
End Class



' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W7105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, sMsg
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7105MA1"
	
	Set lgcTB_50A = New C_TB_50A		' -- 해당서식 클래스 
	Set lgcTB_3_2_2= New C_TB_3_2_2		' -- 해당서식 클래스 
	Set lgcCOMPANY_HISTORY= new C_COMPANY_HISTORY
	
	call lgcCOMPANY_HISTORY.LoadData
	
	If Not lgcTB_50A.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	If Not lgcTB_3_2_2.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 

	'==========================================
	' -- 제3호 법인세과세표준 및 세액조정계산서 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	Do Until lgcTB_50A.EOF 
	
		Select Case lgcTB_50A.W_CD
			Case "01"
				sMsg = "1.자본금"
			Case "02"
				sMsg = "2.자본잉여금"
			Case "03"
				sMsg = "3.이익잉여금"
			Case "04"
				sMsg = "4.자본조정"
			Case "05"
				sMsg = ""
			Case "06"
				sMsg =""
			Case "07"
				sMsg = ""
			Case "08"
				sMsg = ""
			Case "09"
				sMsg = ""
			Case "10"
				sMsg = ""
			Case "11"
				sMsg = ""
			Case "12"
				sMsg =""
			Case "13"
				sMsg = ""
			Case "20"
				sMsg = "5.계(I)"
			Case "21"
				sMsg = "6.자분금과적립금계산서(을)계(II)"
			Case "22"
				sMsg = "7.법인세"
			Case "23"
				sMsg = "8.주민세"
			Case "30"
				sMsg = "9.계(III)"
			Case "31"
				sMsg = "10.차가감계(I+II-III)"			
		End Select
	
		Select Case lgcTB_50A.W_CD
			Case "16","17"
				sHTFBody = sHTFBody & UNIChar(lgcTB_50A.W1, 30)	' Null 허용 
		End Select
		
		If Not ChkNotNull(lgcTB_50A.W2, sMsg & "_기초잔액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W2, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W3, sMsg & "_당기중증감_감소") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W3, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W4, sMsg & "_당기중증감_증가") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W4, 15, 0)
		
		If Not ChkNotNull(lgcTB_50A.W5, sMsg & "_기말잔액") Then blnError = True	
		
		'200703 add
		
		
		
		Select Case lgcTB_50A.W_CD
			Case "01" '자본금 
		
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W1) then
					
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W1) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"표준대차대조표-자본금", sMsg & "_기말잔액"))
				
				end if
				
			Case "02" '자본잉여금 
			
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W2) then
					Call SaveHTFError(lgsPGM_ID,cDbl(lgcTB_3_2_2.W2) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"표준대차대조표-자본잉여금", sMsg & "_기말잔액"))
				end if
				
			Case "14" '이익잉영금 
			
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W3) then
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W3) & "<>" &cDbl(lgcTB_50A.W5)  , UNIGetMesg(TYPE_CHK_NOT_EQUAL,"표준대차대조표-이익잉영금", sMsg & "_기말잔액"))
				end if
			
			Case "15" '자본조정 
				if cDbl(lgcTB_50A.W5)<> cDbl(lgcTB_3_2_2.W4) then
					Call SaveHTFError(lgsPGM_ID, cDbl(lgcTB_3_2_2.W4) & "<>" &cDbl(lgcTB_50A.W5), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"표준대차대조표-자본조정", sMsg & "_기말잔액"))
				end if
				
				
			case else 
			
		end Select 

		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50A.W5, 15, 0)

		lgcTB_50A.MoveNext 
	Loop
	
	
	sHTFBody = sHTFBody & UNIChar("", 64)	' -- 공란 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_50A = Nothing	' -- 메모리해제 
	
End Function


%>
