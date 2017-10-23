<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제2호 농어촌특별세 과세표준 및 세액신고 
'*  3. Program ID           : W8113MA1
'*  4. Program Name         : W8113MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_2

Set lgcTB_2 = Nothing	' -- 초기화 

Class C_TB_2
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W7
	Dim W8
	Dim W9
	Dim W10_1
	Dim W10_2
			
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs1
			 
		On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		LoadData = False
			 
		PrintLog "LoadData IS RUNNING: "
			 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		lgStrSQL = ""
		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W1				= oRs1("W1")
		W2				= oRs1("W2")
		W3				= oRs1("W3")
		W4				= oRs1("W4")
		W5				= oRs1("W5")
		W6				= oRs1("W6")
		W7				= oRs1("W7")
		W8				= oRs1("W8")
		W9				= oRs1("W9")
		W10_1			= oRs1("W10_1")
		W10_2			= oRs1("W10_2")
		
		Call SubCloseRs(oRs1)	
		
		LoadData = True
	End Function

	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub	

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_2	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W8113MA1
	Dim A105
	Dim A101
	
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8113MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8113MA1"
	
	Set lgcTB_2 = New C_TB_2		' -- 해당서식 클래스 
	
	If Not lgcTB_2.LoadData	Then Exit Function		' -- 제1호 서식 로드 
	
	Set cDataExists = new TYPE_DATA_EXIST_W8113MA1	
	'==========================================
	' -- 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	

	
	If ChkNotNull(lgcTB_2.W1, "과세표준")  Then ' -- 데이타존재시 검증식 
	
			
			' -- 제12호농어촌특별세과세표준및세액조정계산서(A105)
			Set cDataExists.A105  = new C_TB_12	' -- W8111MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8111MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A105.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A105.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A105.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제12호농어촌특별세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				'일반법인 : 농어촌특별세과세표준및세액조정계산서(A105)의 항목(8)  일반법인 과세표준금액_소계또는 
				'조합법인 : 농어촌특별세과세표준및세액조정계산서(A105)의 항목(12) 조합법인 과세표준금액_소계와 일치치하지 않으면 오류 
				If (UNICDbl(lgcTB_2.W1, 0) <> UNICDbl(cDataExists.A105.w8_Amt, 0))  And  (UNICDbl(lgcTB_2.W1, 0) <> UNICDbl(cDataExists.A105.w12_Amt, 0) )Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준(7)","제12호농어촌특별세과세표준및세액조정계산서(A105) 항목(8)  과세표준금액_소계"))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A105 = Nothing
		
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W1, 15, 0)
	

	If ChkNotNull(lgcTB_2.W2, "산출세액")  Then ' -- 데이타존재시 검증식 
	
			
			' -- 제12호농어촌특별세과세표준및세액조정계산서(A105)
			Set cDataExists.A105  = new C_TB_12	' -- W8111MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A105.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A105.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A105.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제12호농어촌특별세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				'일반법인 : 농어촌특별세과세표준및세액조정계산서(A105)의 항목(8)  일반법인 산출세액_소계또는 
				'조합법인 : 농어촌특별세과세표준및세액조정계산서(A105)의 항목(12) 조합법인 산출세액_소계와 일치치하지 않으면 오류 
				If (UNICDbl(lgcTB_2.W2, 0) <> UNICDbl(cDataExists.A105.w8_Tax, 0))  And (UNICDbl(lgcTB_2.W2, 0) <> UNICDbl(cDataExists.A105.w12_Tax, 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준(7)","항목(8)  산출세액_소계"))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A105 = Nothing
	
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W2, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W3, "가산세액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W3, 15, 0)
	
	'총부담세액(10) : (8)산출세액 + (9)가산세액 
	
	If ChkNotNull(lgcTB_2.W4, "총부담세액") Then
	    if UNICDbl(lgcTB_2.W4, 0) <> UNICDbl(lgcTB_2.W2,0) + UNICDbl(lgcTB_2.W3, 0) then
	       	blnError = True
		
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "총부담세액","산출세액 + 가산세액"))
	    end if
	else
	   blnError = True	
	end if  

	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W4, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W5, "기납부세액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W5, 15, 0)
	
	
	'차감납부할세액(12) : (10)총부담세액 - (11)기납부세액 
	If ChkNotNull(lgcTB_2.W6, "차감납부할세액") Then
	    if UNICDbl(lgcTB_2.W6, 0) <> UNICDbl(lgcTB_2.W4,0) - UNICDbl(lgcTB_2.W5, 0) then
	       	blnError = True
		
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부할세액","총부담세액 - 기납부세액"))
	    end if
	else
	   blnError = True	
	end if  	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W6, 15, 0)
	
	
	'차감납부세액(14)  : (12)차감납부할세액 - (13)분납할세액 
	If Not ChkNotNull(lgcTB_2.W7, "분납할세액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W7, 15, 0)
	

	If ChkNotNull(lgcTB_2.W8, "차감납부세액") Then
	    if UNICDbl(lgcTB_2.W8, 0) <> UNICDbl(lgcTB_2.W6,0) - UNICDbl(lgcTB_2.W7, 0) then
	       	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W8, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부세액","차감납부할세액 - 분납할세액"))
	    end if
	else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W8, 15, 0)
	

	'충당후납부세액(15) : (14)차감납부세액  - (16)충당할농어촌특별세 
	If ChkNotNull(lgcTB_2.W9, "충당후납부세액") Then
	    if UNICDbl(lgcTB_2.W9, 0) <> UNICDbl(lgcTB_2.W8,0) - UNICDbl(lgcTB_2.W10_2, 0) then
	       	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_2.W9, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "충당후납부세액","차감납부세액 - 충당할농어촌특별세"))
	    end if
	else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W9, 15, 0)
	
	If Not ChkNotNull(lgcTB_2.W10_1, "국세환급금충당신청_환급법인세") Then blnError = True	
	
	'- 음수는 오류 
	'- 법인세가 환급인 경우 입력가능 
	'- 입력된경우 ZERO보다 크고 법인세과세표준및세액조정계산서(A101)의 
	'  코드(46)차감납부할세액계(코드(46) >= 0 오류)와 절대값이 일치해야 함 
    '  (예: 코드(46)이 -100,000 이면 환급법인세는 ZERO이거나 100,000 으로 입력되어야 함)
	'- 법인세과세표준및세액조정계산서(A101)의 코드(46)차감납부할세액계 >= 0 입력불가 
	
	

	If UNICDbl(lgcTB_2.W10_1, 0) >= 0 Then	' -- 환급법인세 음수 오류체크 
	   	if  UNICDbl(lgcTB_2.W10_1, 0) > 0 then	
				' -- 제3호법인세과세표준및세액조정계산서(A101)
				Set cDataExists.A101  = new C_TB_3	' -- W8111MA1_HTF.asp 에 정의됨 
			
				' -- 추가 조회조건을 읽어온다.
				Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
				cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
				cDataExists.A101.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
				If Not cDataExists.A101.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "제3호법인세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
				Else
		
					If  (abs(UNICDbl(lgcTB_2.W10_1, 0) <> UNICDbl(cDataExists.A101.w46, 0)) Or UNICDbl(cDataExists.A101.w46, 0) >= 0)Then
						blnError = True
					
						Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "환급법인세","법인세과세표준및세액조정계산서(A101)의 코드(46)차감납부할세액계(코드(46) >= 0 오류)와 절대값"))
					End If
				End If
		
				' -- 사용한 클래스 메모리 해제 
				Set cDataExists.A101 = Nothing
		end if		
	
	
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_1, UNIGetMesg(TYPE_CHK_ZERO_OVER, "국세환급금충당신청_환급법인세",""))
	End If
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W10_1, 15, 0)
	
	
		
	If Not ChkNotNull(lgcTB_2.W10_2, "국세환급금충당신청_충당할농어촌특별세") Then blnError = True	
	'- 음수는 오류 
	'- 법인세가 환급인 경우 입력가능 
	'- ZERO 이거나 입력된경우 ZERO보다 크고 법인세과세표준및 세액조정계산서(A101)
	'  코드(46)차감납부할세액계보다 절대값이 작거나 같아야 함 
	'- 입력된경우 ZERO보다 크고 항목(14)차감납부세액보다 작거나 같아야 함 
	
	If UNICDbl(lgcTB_2.W10_2, 0) >= 0 Then	' -- 환급법인세 음수 오류체크 
	   		
			' -- 제3호법인세과세표준및세액조정계산서(A101)
			Set cDataExists.A101  = new C_TB_3	' -- W8111MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8113MA1("0",iKey1, iKey2, iKey3)   
			
			cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A101.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A101.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제3호법인세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				
				If UNICDbl(lgcTB_2.W10_2, 0) > abs(UNICDbl(cDataExists.A101.w46, 0)) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_2, UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "환급법인세","법인세과세표준및 세액조정계산서(A101)(46)차감납부할세액계보다 절대값이 "))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A101 = Nothing
	
	
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_2.W10_2, UNIGetMesg(TYPE_CHK_ZERO_OVER, "국세환급금충당신청_충당할농어촌특별세",""))
	End If
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_2.W10_2, 15, 0)


	sHTFBody = sHTFBody & UNIChar("", 29)	' -- 공란 
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	' -- 파일에 기록한다.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If
	
	'Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_2 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 

	  Case "O" '-- 외부 참조 금액 
	
			
	End Select
	PrintLog "SubMakeSQLStatements_W8113MA1 : " & lgStrSQL
End Sub

%>
