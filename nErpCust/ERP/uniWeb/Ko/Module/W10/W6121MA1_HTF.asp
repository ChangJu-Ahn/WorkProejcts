<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  재8호 공제감면세액계산서(3)
'*  3. Program ID           : W6121MA1
'*  4. Program Name         : W6121MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_8_3

Set lgcTB_8_3 = Nothing ' -- 초기화 

Class C_TB_8_3
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	Dim SELECT_SQL
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs3		' -- 멀티로우 데이타는 지역변수로 선언한다.

	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True : blnData3 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' --서식을 읽어온다.
		Call SubMakeSQLStatements("A",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("B",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		
		Call SubMakeSQLStatements("C",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
		
		
		
		If blnData1 = False And blnData2 = False And   blnData3 = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Call MoveFirst(pType)
		Select Case pType
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
			Case 3
				lgoRs3.Find pWhereSQL	
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
			Case 3
				lgoRs3.Filter = pWhereSQL	
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
			Case 3
				EOF = lgoRs3.EOF	
		End Select
	End Function
	
	Function MoveFirst(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
			Case 3
				lgoRs3.MoveFirst	
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
			Case 3
				lgoRs3.MoveNext	
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
		On Error Resume Next
		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
					If Err Then PrintLog "pFieldNm=" & pFieldNm : Reponse.End
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
				End If
			Case 3
				If Not lgoRs3.EOF Then
					GetData = lgoRs3(pFieldNm)
				End If	
		End Select
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)		
	End Sub

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf	
				lgStrSQL = lgStrSQL & " FROM TB_8_3_A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				lgStrSQL = lgStrSQL & "  ORDER BY  CAST(W101 as INT)  ASC " & vbcrlf

	      Case "B"
				If WHERE_SQL <> "" Then
				
					lgStrSQL = ""
					lgStrSQL = lgStrSQL & " SELECT  "
					lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
					lgStrSQL = lgStrSQL & " FROM TB_8_3_B	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
					lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				Else
				
					lgStrSQL =			  "SELECT "
					lgStrSQL = lgStrSQL & "  SEQ_NO   , A.W105, A.W105_NM, A.W106, A.C_W107, A.C_W108, A.C_W109, A.C_W110, A.C_W111, A.C_W112, A.C_W113, A.C_W114, A.C_W115, A.C_W116, A.C_W117, A.C_W118 "
					lgStrSQL = lgStrSQL & " FROM TB_8_3_B A WITH (NOLOCK) "
					lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & "  union all"
					lgStrSQL = lgStrSQL & "	 SELECT 9999999 SEQ_NO ,  '99999' , '계','',Sum(A.C_W107), Sum(A.C_W108), Sum(A.C_W109), Sum(A.C_W110), Sum(A.C_W111), Sum(A.C_W112), Sum(A.C_W113), Sum(A.C_W114), Sum(A.C_W115), Sum(A.C_W116),Sum(A.C_W117), Sum(A.C_W118 )"
					lgStrSQL = lgStrSQL & "  FROM TB_8_3_B A WITH (NOLOCK) "
					lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
					lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & "  ORDER BY  W105 ,  W106  ASC " & vbcrlf
					
					
				End If	

				
			Case "C"	
				 lgStrSQL = ""
				 lgStrSQL = lgStrSQL &  " SELECT Sum(A.C_W107) C_W107, Sum(A.C_W108) C_W108, Sum(A.C_W116) C_W116,Sum( A.C_W118)  C_W118  " & vbCrLf
				 lgStrSQL = lgStrSQL &  " FROM TB_8_3_B A WITH (NOLOCK) " & vbCrLf	' 서식3호 
				 lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
				 lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				 lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				 If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6121MA1
	Dim A103
	Dim A181
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6121MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo , dblSum , dblCode30
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6121MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6121MA1"

	Set lgcTB_8_3 = New C_TB_8_3		' -- 해당서식 클래스 
	
	If Not lgcTB_8_3.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W6121MA1

	'==========================================
	' --  재8호 공제감면세액계산서(3) 오류검증 
	' -- 1. 매출및매입거래등 
	
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	dblSum = 0
	dblCode30 = 0	
	
	lgcTB_8_3.Find 1, "W_Code='01'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='16'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='02'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='04'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='05'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='06'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='07'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='08'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='09'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='10'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='11'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='15'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='12'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='17'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='18'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='13'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "W101_NM"), 40)      '구분 
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='30'"
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='19'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	lgcTB_8_3.Find 1, "W_Code='20'"
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)      '계산내역 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
	
	sHTFBody = sHTFBody & UNIChar("", 29) & vbCrLf	' -- 공란 
	
	' -- 개정되어 파일생성 순서가 바뀜 
	' ---------------------------------------------------------

	lgcTB_8_3.Find 1, "W_Code='01'"
	
	Do Until lgcTB_8_3.EOF(1) 
		SELECT CASE  lgcTB_8_3.GetData(1, "W_CODE")
		
	
		CASE   "13" , "14"
		
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "W101_NM"), 40)      '구분 
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)          '계산내역 
		    	'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0) '공제대상세액 
		
		       dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '각 공제대상액의 합 
		
		
		CASE   "30" 
		
		      '공제대상 합계 
		
			If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W104"), lgcTB_8_3.GetData(1, "W101_NM") & "_공제대상세액") Then blnError = True	
			'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0)
		
		    dblCode30 = unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)  '공제대상 합계 

' -- 2006.03.27 개정추가 
'    세액공제조정명세서(3)(A149)서식의 코드04. 연구인력개발비세액공제의 항목(104)공제대상세액이 0보다 클 경우 
'     연구및인력개발비발생명세서(A181)의 서식이 없으면 오류 
'    -> 검증 추가 
		Case "04"

			If UNICDbl(lgcTB_8_3.GetData(1, "C_W104"), 0) > 0 Then	' -- 대상세액 
			  '- 코드(32)의 공제세액이"0"보다 큰 경우 연구및인력개발비발생명세서(A181)의 서식유무 검증 

				Set cDataExists.A181 = new C_TB_JT3	' -- W6111MA1_HTF.asp 에 정의됨 
								
				' -- 추가 조회조건을 읽어온다.
				cDataExists.A181.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
				cDataExists.A181.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
				If Not cDataExists.A181.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "", lgcTB_8_3.GetData(1, "W101_NM") & "_대상세액이 '0'보다 큰 경우 연구및인력개발비발생명세서(A181) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
						
				End If
							
				' -- 사용한 클래스 메모리 해제 
				Set cDataExists.A181 = Nothing							
			End If	

			dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '각 공제대상액의 합 
		CASE ELSE
		      	'If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W103"), lgcTB_8_3.GetData(1, "W101_NM") & "_계산내역") Then blnError = True	' -- Null 허용 
				'sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(1, "C_W103"), 30)
		
				If Not ChkNotNull(lgcTB_8_3.GetData(1, "C_W104"), lgcTB_8_3.GetData(1, "W101_NM") & "_공제대상세액") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(1, "C_W104"), 15, 0)
		
		       dblSum  = dblSum + unicdbl(lgcTB_8_3.GetData(1, "C_W104") ,0)    '각 공제대상액의 합 
		      
		END SELECT		
				
		
		Call lgcTB_8_3.MoveNext(1)	' -- 1번 레코드셋 
	Loop
	lgcTB_8_3.WHERE_SQL = ""
	      
	
	if dblCode30 <> dblSum then
	    Call SaveHTFError(lgsPGM_ID, dblCode30 & " <> " & dblSum, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"공제대상 합계", "각 공제대상액의 합"))
	    blnError = True	
	End if
	
	if Not lgcTB_8_3.EOF(3)  then
 
		if dblCode30 <> unicdbl(lgcTB_8_3.GetData(3, "C_W107"),0) then
		    Call SaveHTFError(lgsPGM_ID, dblCode30, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"공제대상 합계", "항목(107)요공제세액_당기분의 합계"))
		    blnError = True	
		End if
	End if
	

	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	' -- 2. 당기및이월액계산 
	iSeqNo = 1

	Do Until lgcTB_8_3.EOF(2) 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용			
		
		If UNICDbl(lgcTB_8_3.GetData(2, "SEQ_NO"), 0) <> 9999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "W106"), "사업연도") Then blnError = True	' 합계행 외엔 사업연도 필수체크 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(2, "SEQ_NO"), 6)
		End If
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "W105"), "구분") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_8_3.GetData(2, "W105"), 2)
		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_8_3.GetData(2, "W106"))	' 사업연도 

		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W107"), "당기분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W107"), 15, 0)
		
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W108"), "이월분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W108"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W109"), "당기공제대상세액_당기분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W109"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W110"), "당기공제대상세액_1차") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W110"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W111"), "당기공제대상세액_2차") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W111"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W112"), "당기공제대상세액_3차") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W112"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W113"), "당기공제대상세액_4차") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W113"), 15, 0)
		
		If  ChkNotNull(lgcTB_8_3.GetData(2, "C_W114"), "당기공제대상세액_합계") Then 
		    if unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W109"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W111"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W112"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W113"),0) then
		       Call SaveHTFError(lgsPGM_ID, unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") & "당기공제대상세액_합계", "항목(109)+항목(110)+항목(111)+항목(112)+항목(113)"))
		       blnError = True	
		    End if
		  
		  
		Else
		    blnError = True	
		End if
		  sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W114"), 15, 0)     
		 
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W115"), "최조한세적용에 따른 미공제액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W115"), 15, 0)
		
		If  ChkNotNull(lgcTB_8_3.GetData(2, "C_W116"), "공제세액") Then 
		    if unicdbl(lgcTB_8_3.GetData(2, "C_W116"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W114"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W115"),0)  then
		       Call SaveHTFError(lgsPGM_ID, lgcTB_8_3.GetData(2, "C_W116"), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") &"공제세액", "항목(114)-항목(115)"))
		       blnError = True	
		    End if
		Else
		    blnError = True	
		End if    
		  
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W116"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W117"), "소멸") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W117"), 15, 0)
		
		If Not ChkNotNull(lgcTB_8_3.GetData(2, "C_W118"), "이월액") Then blnError = True	
		     if unicdbl(lgcTB_8_3.GetData(2, "C_W118"),0) <> unicdbl(lgcTB_8_3.GetData(2, "C_W107"),0) + unicdbl(lgcTB_8_3.GetData(2, "C_W108"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W116"),0) - unicdbl(lgcTB_8_3.GetData(2, "C_W117"),0) then
		       Call SaveHTFError(lgsPGM_ID, unicdbl(lgcTB_8_3.GetData(2, "C_W118"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_8_3.GetData(2, "W105_Nm") &"이월액", "항목(107)+항목(108) -항목(116)-항목(117) "))
		       blnError = True	
		 End if
		sHTFBody = sHTFBody & UNINumeric(lgcTB_8_3.GetData(2, "C_W118"), 15, 0)

		sHTFBody = sHTFBody & UNIChar("", 38) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_8_3.MoveNext(2)	' -- 2번 레코드셋 
	Loop

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_8_3 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6121MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W6121MA1 : " & lgStrSQL
End Sub
%>
