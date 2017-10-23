<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제15호 과목별소득금액조정명세서 
'*  3. Program ID           : W5103MA1
'*  4. Program Name         : W5103MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_15

Set lgcTB_15 = Nothing ' -- 초기화 

Class C_TB_15
	' -- 테이블의 컬럼변수 
	
	Dim SELECT_SQL
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.

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
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm)
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
				
	            If WHERE_SQL = "" Then 
					lgStrSQL = lgStrSQL & " , ( SELECT ITEM_NM FROM TB_ADJUST_ITEM WITH (NOLOCK)  WHERE ITEM_CD = A.W1 ) W1_NM " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & " , '' W1_NM  " & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_15	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W5103MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W5103MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W5103MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W5103MA1"

	Set lgcTB_15 = New C_TB_15		' -- 해당서식 클래스 
	
	If Not lgcTB_15.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W5103MA1

	'==========================================
	' -- 제15호 소득금액조정합계표 오류검증 
	iSeqNo = 1	: sHTFBody = ""
	
	Do Until lgcTB_15.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W_TYPE"), 1)
		
		If UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_15.GetData("W3"), lgcTB_15.GetData("W1_NM") & " 처분코드") Then blnError = True	
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("SEQ_NO"), 6)
			iSeqNo = 1	' -- W_TYPE 으로 인해 초기화 
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W1_NM"), 40)
		
		If  Not ChkMinusAmt(lgcTB_15.GetData("W2"),"음수체크") Then blnError = True	' -- 음수체크 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_15.GetData("W2"), 15, 0)
 
		If lgcTB_15.GetData("W_TYPE") = "1" and UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999  Then
			If Not ChkBoundary("100,200,300,400,500,600", lgcTB_15.GetData("W3"), lgcTB_15.GetData("W_TYPE") & " 처분코드") Then blnError = True
		ElseIf lgcTB_15.GetData("W_TYPE") = "2" and UNICDbl(lgcTB_15.GetData("SEQ_NO"), 0) <> 999999 Then
			If Not ChkBoundary("100,200", lgcTB_15.GetData("W3"), lgcTB_15.GetData("W1_NM") & " 처분코드") Then blnError = True
			 
		End If
			
		sHTFBody = sHTFBody & UNIChar(lgcTB_15.GetData("W3"), 4)
	
		If lgcTB_15.GetData("W_TYPE") = "1" And lgcTB_15.GetData("W3") = "400" Then sT1_400SUM = sT1_400SUM + UNICDbl(lgcTB_15.GetData("W2"), 0)
		If lgcTB_15.GetData("W_TYPE") = "2" And lgcTB_15.GetData("W3") = "100" Then sT2_100SUM = sT2_100SUM + UNICDbl(lgcTB_15.GetData("W2"), 0)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		lgcTB_15.MoveNext 
	Loop
	
	' 익금산입및손금불산입의 코드 '400'의 합 또는 손금산입및익금불산입의 코드 '100'의 합 금액이 '0'이 아니면 자본금과적립금조정명세서(을)(A103) 서식 필수 입력 
	If sT1_400SUM > 0 Or sT2_100SUM > 0 Then
		' -- 제50호 자본금과 적립금조정명세서(을)
		Set cDataExists.A103 = new C_TB_50A	' -- W7105MA1_HTF.asp 에 정의됨 
			
		' -- 추가 조회조건을 읽어온다.
		cDataExists.A103.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A103.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
		If Not cDataExists.A103.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "익금산입및손금불산입의 코드 '400'의 합 또는 손금산입및익금불산입의 코드 '100'의 합 금액이 '0'이 아니면 자본금과적립금조정명세서(을)(A103) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		End If
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A103 = Nothing
	End If

	' ----------- 
	'Call SubCloseRs(oRs2)

	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_15 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W5103MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W5103MA1 : " & lgStrSQL
End Sub
%>
