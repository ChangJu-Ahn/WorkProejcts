<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제12호 농특세 과세표준및 세액조정계산서 
'*  3. Program ID           : W8111MA1
'*  4. Program Name         : W8111MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_12

Set lgcTB_12 = Nothing	' -- 초기화 

Class C_TB_12
	' -- 테이블의 컬럼변수 
	Dim W5_AMT
	Dim W5_RATE
	Dim W5_RATE_VAL
	Dim W5_TAX
	Dim W6
	Dim W6_AMT
	Dim W6_RATE
	Dim W6_RATE_VAL
	Dim W6_TAX
	Dim W7
	Dim W7_AMT
	Dim W7_RATE
	Dim W7_RATE_VAL
	Dim W7_TAX
	Dim W8_AMT
	Dim W8_TAX
	Dim W10_AMT
	Dim W10_RATE
	Dim W10_RATE_VAL
	Dim W10_TAX
	Dim W11
	Dim W11_AMT
	Dim W11_RATE
	Dim W11_RATE_VAL
	Dim W11_TAX
	Dim W12_AMT
	Dim W12_TAX
			
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
		
		W5_AMT			= oRs1("W5_AMT")
		W5_RATE			= oRs1("W5_RATE")
		W5_RATE_VAL		= oRs1("W5_RATE_VAL")
		W5_TAX			= oRs1("W5_TAX")
		W6				= oRs1("W6")
		W6_AMT			= oRs1("W6_AMT")
		W6_RATE			= oRs1("W6_RATE")
		W6_RATE_VAL		= oRs1("W6_RATE_VAL")
		W6_TAX			= oRs1("W6_TAX")
		W7				= oRs1("W7")
		W7_AMT			= oRs1("W7_AMT")
		W7_RATE			= oRs1("W7_RATE")
		W7_RATE_VAL		= oRs1("W7_RATE_VAL")
		W7_TAX			= oRs1("W7_TAX")
		W8_AMT			= oRs1("W8_AMT")
		W8_TAX			= oRs1("W8_TAX")
		W10_AMT			= oRs1("W10_AMT")
		W10_RATE		= oRs1("W10_RATE")
		W10_RATE_VAL	= oRs1("W10_RATE_VAL")
		W10_TAX			= oRs1("W10_TAX")
		W11				= oRs1("W11")
		W11_AMT			= oRs1("W11_AMT")
		W11_RATE		= oRs1("W11_RATE")
		W11_RATE_VAL	= oRs1("W11_RATE_VAL")
		W11_TAX			= oRs1("W11_TAX")
		W12_AMT			= oRs1("W12_AMT")
		W12_TAX			= oRs1("W12_TAX")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_12	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W8111MA1
	Dim A151

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8111MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8111MA1"
	
	Set lgcTB_12 = New C_TB_12		' -- 해당서식 클래스 
	
	If Not lgcTB_12.LoadData	Then Exit Function		' -- 제12호 서식 로드 
	
	Set cDataExists = new TYPE_DATA_EXIST_W8111MA1
	'==========================================
	' -- 제12호 농특세 과세표준및 세액조정계산서 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	If ChkNotNull(lgcTB_12.W5_AMT, "일반법인_과세표준금액")  Then ' -- 데이타존재시 검증식 
	
			
			' --제13호농어촌특별세과세대상감면세액합계표 
			Set cDataExists.A151  = new C_TB_13	' -- W8109MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8111MA1("A105",iKey1, iKey2, iKey3)   
			
			cDataExists.A151.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A151.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A151.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제13호농어촌특별세과세대상감면세액합계표", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				cDataExists.A151.FIND 1, " W1_CD='10' "
				If  (UNICDbl(lgcTB_12.W5_AMT, 0) <> UNICDbl(cDataExists.A151.GetData(1, "W4"), 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_12.W5_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[일반법인] 과세표준금액","농어촌특별세과세대상감면세액합계표의(A151)의 항목(10)의 금액"))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A151 = Nothing
	
	Else
		blnError = True
	End If

	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_AMT, 15, 0)
	
	If Not ChkNotNull(lgcTB_12.W5_RATE, "일반법인_세율") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_RATE, 5, 2)

	
	If ChkNotNull(lgcTB_12.W5_TAX, "일반법인_세액") Then 
	   if  UNICDbl(lgcTB_12.W5_TAX, 0) <>   Fix((UNICDbl(lgcTB_12.W5_AMT, 0) * UNICDbl(lgcTB_12.W5_RATE_VAL,0))) then
	       blnError = True	
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W5_TAX & " <> " & Int((UNICDbl(lgcTB_12.W5_AMT, 0) * UNICDbl(lgcTB_12.W5_RATE_VAL,0))), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[일반법인]세액","항목(5)의 과세표준금액 Ｘ 세율(" & lgcTB_12.W5_RATE &")"))
	   end if
	else

		blnError = True	
	end if	
	   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W5_TAX, 15, 0)
	
	'기타1구분 
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W6, 20)
	
	
	
	If  ChkNotNull(lgcTB_12.W6_AMT, "일반법인(기타1)_" & lgcTB_12.W6 & "_세액") Then 
	    if UNICDbl(lgcTB_12.W6_AMT, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타1_금액 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_AMT, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W6_RATE, "일반법인(기타1)_" & lgcTB_12.W6 & "_금액") Then 
	    if UNICDbl(lgcTB_12.W6_RATE, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타1_비율 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_RATE, 5, 2)
	
	
	
	If  ChkNotNull(lgcTB_12.W6_TAX, "일반법인(기타1)_" & lgcTB_12.W6 & "_세액") Then 
	    if UNICDbl(lgcTB_12.W6_TAX, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W6_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타1_세액 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W6_TAX, 15, 0)
	
	'기타2구분 
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W7, 20)
	
	
	
	If  ChkNotNull(lgcTB_12.W7_AMT, "일반법인(기타2)_" & lgcTB_12.W7 & "_금액") Then 
	    if UNICDbl(lgcTB_12.W7_AMT, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타2_금액 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_AMT, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W7_RATE, "일반법인(기타2)_" & lgcTB_12.W7 & "_세액") Then 
	    if UNICDbl(lgcTB_12.W7_RATE, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타2_비율 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_RATE, 5, 2)
	
	
	
	If  ChkNotNull(lgcTB_12.W7_TAX, "일반법인(기타1)_" & lgcTB_12.W7 & "_세액") Then 
	    if UNICDbl(lgcTB_12.W7_TAX, 0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W7_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[일반법인]기타1_금액 ",""))
	    
	        blnError = True	
	    end if    
	    
	Else
	   blnError = True	
	end if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W7_TAX, 15, 0)
	
	
	
	If ChkNotNull(lgcTB_12.W8_AMT, "일반법인_소계_과세표준금액") Then 
	   If UNICDbl(lgcTB_12.W8_AMT,  0) <> UNICDbl(lgcTB_12.W5_AMT, 0) + UNICDbl(lgcTB_12.W6_AMT,  0) + UNICDbl(lgcTB_12.W7_AMT,  0) then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[일반법인]소계_과세표준금액","항목(5)의 과세표준금액 + 항목(6)의 과세표준금액 + 항목(7)의 과세표준금액"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W8_AMT, 15, 0)
	
	If ChkNotNull(lgcTB_12.W8_TAX, "일반법인_소계_과세표준금액") Then 
	   If UNICDbl(lgcTB_12.W8_TAX, 0) <> UNICDbl(lgcTB_12.W5_TAX,0) + UNICDbl(lgcTB_12.W6_TAX, 0) + UNICDbl(lgcTB_12.W7_TAX, 0) then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[일반법인]소계_세액","항목(5)의 세액 + 항목(6)의 세액 + 항목(7)의 세액"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W8_TAX, 15, 0)
	
	If  ChkNotNull(lgcTB_12.W10_AMT, "조합법인등_과세표준금액") Then 
	    	' --제13호농어촌특별세과세대상감면세액합계표 
			Set cDataExists.A151  = new C_TB_13	' -- W8109MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8111MA1("A151",iKey1, iKey2, iKey3)   
			
			cDataExists.A151.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A151.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A151.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제13호농어촌특별세과세대상감면세액합계표", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				    
				  
				If  (UNICDbl(lgcTB_12.W10_AMT, 0) <> UNICDbl(cDataExists.A151.GetData(2, "W7"), 0) )Then
					blnError = True
				
					Call SaveHTFError(lgsPGM_ID, lgcTB_12.W10_AMT, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[조합법인등]과세표준금액","농어촌특별세과세대상감면세액합계표(A151)의 조합법인등의 감면세액항목(7)의 합"))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A151 = Nothing
	else
	
	   blnError = True	
	end if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_AMT, 15, 0)
	
	
	
	If Not ChkNotNull(lgcTB_12.W10_RATE, "조합법인등_세율") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_RATE, 5, 2)
	
	If ChkNotNull(lgcTB_12.W10_TAX, "조합법인등_세액") Then 
	   if  UNICDbl(lgcTB_12.W10_TAX,  0) <>  Fix(UNICDbl(lgcTB_12.W10_AMT,0) *  UNICDbl(lgcTB_12.W10_RATE, 2)) then
	       blnError = True	
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W10_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[조합법인등]세액","항목(10)의 과세표준금액 Ｘ 세율(20%) "))
	   end if
	else

		blnError = True	
	end if	
	   	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W10_TAX, 15, 0)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_12.W11, 20)
	

	

    If  ChkNotNull(lgcTB_12.W11_AMT, "조합법인등(기타1)_금액")  then
	      if UNICDbl(lgcTB_12.W11_AMT,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_AMT, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[조합법인등]기타금액세율 ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_AMT, 15, 0)
	
	
	If  ChkNotNull(lgcTB_12.W11_RATE, "조합법인등(기타1)_세율")  then
	    if UNICDbl(lgcTB_12.W11_RATE,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_RATE, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[조합법인등]기타1_세율 ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_RATE, 5, 2)
	

	
	If  ChkNotNull(lgcTB_12.W11_TAX, "조합법인등(기타1)_세액")  then
	    if UNICDbl(lgcTB_12.W11_TAX,  0) <> 0 then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_12.W11_TAX, UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "[조합법인등]기타1_세액 ",""))
	    
	        blnError = True	
	    end if    
	else
	
	  blnError = True	
	end if 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W11_TAX, 15, 0)
	

	If ChkNotNull(lgcTB_12.W12_AMT, "조합법인등_소계_과세표준금액") Then 
	   If UNICDbl(lgcTB_12.W12_AMT,  0) <> UNICDbl(lgcTB_12.W10_AMT, 0) + UNICDbl(lgcTB_12.W11_AMT, 0)  then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W8_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[조합법인등]소계_과세표준금액","항목(10)의 과세표준금액 + 항목(11)의 과세표준금액"))
	   End if
	Else
	   
	   blnError = True	
	End if   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W12_AMT, 15, 0)
	
	
	
	
	If Not ChkNotNull(lgcTB_12.W12_TAX, "조합법인등__세액") Then blnError = True	
	If ChkNotNull(lgcTB_12.W12_TAX, "조합법인등_소계_세액") Then 
	   If UNICDbl(lgcTB_12.W12_TAX, 0) <> UNICDbl(lgcTB_12.W10_TAX, 0) + UNICDbl(lgcTB_12.W11_TAX,  0)  then
	       blnError = True
	      Call SaveHTFError(lgsPGM_ID, lgcTB_12.W12_TAX, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "[조합법인등]소계_세액","항목(10)의 세액 + 항목(11)의 세액"))
	   End if
	Else
	   
	   blnError = True	
	End if  
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_12.W12_TAX, 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 49) 
	
	
	' -- 파일에 기록한다.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If
	
	'Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_12 = Nothing	' -- 메모리해제  <-- W8101MA1_HTF에서 사용함 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 

	  Case "A151" '-- 외부 참조 금액 
	
			lgStrSQL = ""
			'lgStrSQL = lgStrSQL & "	AND  A.W1_CD	= '10' 	" & vbCrLf
	
			
	End Select
	PrintLog "SubMakeSQLStatements_W8111MA1 : " & lgStrSQL
End Sub

%>
