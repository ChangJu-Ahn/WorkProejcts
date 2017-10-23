<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제10호 원천납부세액명세서 
'*  3. Program ID           : W7101MA1
'*  4. Program Name         : W7101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_10A

Set lgcTB_10A = Nothing ' -- 초기화 

Class C_TB_10A
	' -- 테이블의 컬럼변수 
	Dim SEQ_NO
	Dim W1
	Dim W2_1
	Dim W2_2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	
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
			SEQ_NO		= lgoRs1("SEQ_NO")
			W1			= lgoRs1("W1")
			W2_1		= lgoRs1("W2_1")
			W2_2		= lgoRs1("W2_2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
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
				lgStrSQL = lgStrSQL & " FROM TB_10A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W7101MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W7101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7101MA1"

	Set lgcTB_10A = New C_TB_10A		' -- 해당서식 클래스 
	
	If Not lgcTB_10A.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	iSeqNo = 1
	'==========================================
	' -- 제10호 원천납부세액명세서 오류검증 

	Do Until lgcTB_10A.EOF() 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
		If UNICDbl(lgcTB_10A.SEQ_NO, 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_10A.W1 , "적요") Then blnError = True	' 합계행 외엔 필수체크 
			If Not ChkNotNull(UNIRemoveDash(lgcTB_10A.W2_1) , "사업자(주민)등록번호") Then blnError = True
			If Not ChkNotNull(lgcTB_10A.W2_2 , "상호(성명)") Then blnError = True
			If Not ChkNotNull(lgcTB_10A.W3 , "원천징수일") Then blnError = True
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_10A.SEQ_NO, 6)
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_10A.W1, 50)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_10A.W2_1), 13)
		sHTFBody = sHTFBody & UNIChar(lgcTB_10A.W2_2, 60)
		sHTFBody = sHTFBody & UNI8Date(lgcTB_10A.W3)
		
		If Not ChkNotNull(lgcTB_10A.W4, "이자금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_10A.W4, 15, 0)
		If UNICDbl(lgcTB_10A.SEQ_NO, 0) <> 999999 Then		
		   If Not ChkNotNull(lgcTB_10A.W5, "세율") Then blnError = True	
		End if   
		sHTFBody = sHTFBody & UNINumeric(Replace(lgcTB_10A.W5,"%",""), 5, 2)
				
		If Not ChkNotNull(lgcTB_10A.W6, "법인세") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_10A.W6, 15, 0)

				
		sHTFBody = sHTFBody & UNIChar("", 22)	 & vbCrLf	' -- 공란 
		iSeqNo = iSeqNo + 1
		Call lgcTB_10A.MoveNext()	' -- 1번 레코드셋 
	Loop
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)	' 루프도는것은 WirteLine이 아님 
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_10A = Nothing	' -- 메모리해제 
	
End Function


%>
