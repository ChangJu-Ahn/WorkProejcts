<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 조특제11호의5고용증대특별세액공제 
'*  3. Program ID           : W6113MA1
'*  4. Program Name         : W6113MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_JT11_5

Set lgcTB_JT11_5 = Nothing	' -- 초기화 

Class C_TB_JT11_5
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1
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

		lgStrSQL = ""
		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgoRs1,lgStrSQL, "", "") = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		
		LoadData = True
	End Function

	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf

	            If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보의 사업개시일(창립일) 불러온다 
					lgStrSQL = lgStrSQL & " , B.FOUNDATION_DT " & vbCrLf
	            End If
	            	            
				lgStrSQL = lgStrSQL & " FROM TB_JT11_5	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 

				If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보를 조인한다.
					lgStrSQL = lgStrSQL & " INNER JOIN TB_COMPANY_HISTORY B WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf
	            End If
	            				
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6113MA1
	Dim A165
	
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6113MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6113MA1"
	
	Set lgcTB_JT11_5 = New C_TB_JT11_5		' -- 해당서식 클래스 
	
	If Not lgcTB_JT11_5.LoadData	Then Exit Function		' -- 제1호 서식 로드 
	Set cDataExists = new  TYPE_DATA_EXIST_W6113MA1		
	'==========================================
	' -- 조특제11호의5고용증대특별세액공제 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("FOUNDATION_DT"), "창업(합병등)일") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT11_5.GetData("FOUNDATION_DT"))
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W1"), "고용증대세액공제액") Then 
	    '고용증대세액공제액 : 항목(7)고용증대인원수 x 1,000,000원 
	    if unicdbl(lgcTB_JT11_5.GetData("W1"),0) <> Unicdbl(lgcTB_JT11_5.GetData("W2"),0) *  Unicdbl(lgcTB_JT11_5.GetData("W1_RATE_VALUE"),0)    then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "고용증대세액공제액", "항목(7)고용증대인원수 x " & Unicdbl(lgcTB_JT11_5.GetData("W1_RATE_VALUE"),0)))
			blnError = True		
	    end if
	else
		blnError = True		
	end if	
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W1"), 15, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W2"), "고용증대인원수") Then
        '항목(8)당해과세연도상시근로자수 - 항목(9)직전과세연도상시근로자수	     
      
	    if  unicdbl(lgcTB_JT11_5.GetData("W2"),0) <> fix(unicdbl(lgcTB_JT11_5.GetData("W3"),0) - unicdbl(lgcTB_JT11_5.GetData("W4"),0)) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W2"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "고용증대인원수", "항목(8)당해과세연도상시근로자수 - 항목(9)직전과세연도상시근로자수"))
			blnError = True		
	    end if
	   
	else
	 blnError = True		
	End if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W2"), 5, 0)
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("W3"), "당해과세연도상시근로자수") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W3"), 7, 2)
	
	If Not ChkNotNull(lgcTB_JT11_5.GetData("W4"), "직전과세연도상시근로자수") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W4"), 7, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W5"), "고용유지세액공제액") Then 
	
	    '항목(11)고용유지인원수 x 500,000원 
	    if unicdbl(lgcTB_JT11_5.GetData("W5"),0) <> fix(unicdbl(lgcTB_JT11_5.GetData("W6"),0) *  Unicdbl(lgcTB_JT11_5.GetData("W5_RATE_VALUE"),0)) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "고용유지세액공제액", "고용유지인원수  x " & Unicdbl(lgcTB_JT11_5.GetData("W5_RATE_VALUE"),0)))
	       blnError = True		
	    end if 
	Else
	    blnError = True		
	End if     
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W5"), 15, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W6"), "고용유지인원수") Then
	  
	  '고용유지인원수 : [(고용유지제도 시행일 전1월간상시근로자1인당 1인평균근로시간 
	  '				-고용유지제도시행일 후1월간상시근로자1인당 1인평균근로시간)
	  '				/ 고용유지제도 시행일 전1월간상시근로자1인당 1인평균근로시간]
	  '				x 직전과세연도상시근로자수  
      '              (소수점미만절사)
          if unicdbl(lgcTB_JT11_5.GetData("W8"),0)  <> 0 then
				if  unicdbl(lgcTB_JT11_5.GetData("W6"),0) <> fix(((unicdbl(lgcTB_JT11_5.GetData("W7"),0) - unicdbl(lgcTB_JT11_5.GetData("W8"),0))/unicdbl(lgcTB_JT11_5.GetData("W7"),0)) * unicdbl(lgcTB_JT11_5.GetData("W9"),0)) then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "고용유지인원수", "[(고용유지제도 시행일 전1월간상시근로자1인당 1인평균근로시-고용유지제도시행일 후1월간상시근로자1인당 1인평균근로시간)/ 고용유지제도 시행일 전1월간상시근로자1인당 1인평균근로시간]x 직전과세연도상시근로자수"))
				    blnError = True		
				end if
		  End if		

	   
	Else
	  blnError = True		
	end if
	
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W6"), 5, 0)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W7"), "고용유지제도 시행일전 1월간 상시근로자 1인당 1인평균 근로시간") Then 
	    '고용유지제도 시행일 전1월간상시근로자1인당 1인평균근로시간 
		'- 24시간 초과하면 오류 
		'- 소수점 2자리 미만 절사 
		if unicdbl(lgcTB_JT11_5.GetData("W7"),0) > 24 then 
		 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W7"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "고용유지제도 시행일전 1월간 상시근로자 1인당 1인평균 근로시간", "24"))
		     blnError = True	
		end if
	Else
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W7"), 5, 2)

	If  ChkNotNull(lgcTB_JT11_5.GetData("W8"), "고용유지제도시행일후 1월간 상시근로자 1인당 1인평균근로시간") Then 
	    '고용유지제도시행일후 1월간 상시근로자 1인당 1인평균근로시간 
		'- 24시간 초과하면 오류 
		'- 소수점 2자리 미만 절사 
		if unicdbl(lgcTB_JT11_5.GetData("W8"),0) > 24 then 
		 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W8"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "고용유지제도시행일후 1월간 상시근로자 1인당 1인평균근로시간", "24"))
		     blnError = True	
		end if
	Else
		blnError = True		
	End if	
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W8"), 5, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W9"), "직전과세연도상시근로자수") Then 
	    If unicdbl(lgcTB_JT11_5.GetData("W9"),0 ) <> unicdbl(lgcTB_JT11_5.GetData("W4"),0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W9"), UNIGetMesg(TYPE_CHK_NOT_EQUAL,"직전과세연도상시근로자수", "항목(9)직전과세연도상시근로자수"))
	       blnError = True		
	    End if   
	Else
	 
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W9"), 7, 2)
	
	If  ChkNotNull(lgcTB_JT11_5.GetData("W10"), "세액공제액 계") Then 
	    ' 항목(6)고용증대세액공제액 + 항목(10)고용유지세액공제액 
		'- 세액공제신청서(A165)의 코드(91) 고용증대 특별세액공제 항목(11)대상세액과 일치 
		'(세액공제액 계가 “0”보다 큰 경우 반드시 입력)
		
		Set cDataExists.A165  = new C_TB_JT1	' -- W6103MA1_HTF.asp 에 정의됨 
		Call SubMakeSQLStatements_W6113MA1("A165",iKey1, iKey2, iKey3)   
		cDataExists.A165.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A165.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 

			 If Not cDataExists.A165.LoadData then
			    blnError = True
				Call SaveHTFError(lgsPGM_ID, "조특 제 1호  세액공제신청서(A165)", TYPE_DATA_NOT_FOUND)	
			else
	
			  	if UNICDBL(lgcTB_JT11_5.GetData("W10"),  0) <> UNICDbl(cDataExists.A165.GetData("W5"), 0) then

			  	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT11_5.GetData("W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제세액(13) "," 세액공제신청서(A165)의 코드(91) 고용증대 특별세액공제 항목(11)대상세액"))
					blnError = True
					
			  	end if
			  	
			  	
			End if    
			Set cDataExists.A165 = Nothing	 
		

	Else
		blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT11_5.GetData("W10"), 15, 0)
		

	sHTFBody = sHTFBody & UNIChar("", 10)	' -- 공란 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_JT11_5 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	   Case "A165" '-- 외부 참조 SQL	
	      ' 세액공제신청서(A165)의 코드(75) 기업의어음제도개선을위한 세액공제 항목(11) 대상세액과 일치 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " and A.W3 = '91'" & vbCrLf

	
	End Select
				Response.Write lgStrSQL
	PrintLog "SubMakeSQLStatements_W6113MA1 : " & lgStrSQL
End Sub

%>
