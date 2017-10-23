<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제47호 주요계정명세서(갑)
'*  3. Program ID           : W9101MA1
'*  4. Program Name         : W9101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_47A

Set lgcTB_47A = Nothing ' -- 초기화 

Class C_TB_47A
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W2_CD
	Dim W3
	Dim W4
	Dim W5
	
	Dim W124
	Dim W125
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs3
				 
		On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' --서식을 읽어온다.
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

		' -- 서식을 읽어온다.
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3) 
		                                      '☜ : Make sql statements
		If   FncOpenRs("R",lgObjConn,oRs3,lgStrSQL, "", "") = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		'Response.End 'zzz
		W124		= oRs3("W124")
		W125		= oRs3("W125")
		

		
		Call SubCloseRs(oRs3)
		
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
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W2_CD		= lgoRs1("W2_CD")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
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
				lgStrSQL = lgStrSQL & " FROM TB_47A1 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				'lgStrSQL = lgStrSQL & "		AND A.W2_CD<>'75' " & pCode3 	' 200703 TEMP
				
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_47A2 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

				
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9101MA1
	Dim A129
	DIM A115
	Dim A101
	Dim A137
	Dim A142
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists,sTmp2
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9101MA1"

	Set lgcTB_47A = New C_TB_47A		' -- 해당서식 클래스 
	
	If Not lgcTB_47A.LoadData Then Exit Function			
	
	Set cDataExists = new TYPE_DATA_EXIST_W9101MA1
	'==========================================
	' -- 제47호 주요계정명세서(갑) 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
 
	Do Until lgcTB_47A.EOF 
	

		Select Case lgcTB_47A.W2_CD 
		

		   	Case "41"
		   	   '코드(41)당연손금기부금의 항목(3)회사계상금액	- 기부금명세서(A129)의 법정기부금(10), 정치자금(20), 문화진흥(60)의 합과 일치 
			   '(코드(41)의 항목(3)의 금액이 “0”보다 큰 경우 A129 반드시 입력)
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
			    
		 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_회사계상금액(41) '0'보다 큰 경우 기부금명세서(A129) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '10' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '20' " 
							     sTmp = sTmp +  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '60' " 
							     sTmp = sTmp +  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(41)당연손금기부금의 항목(3)회사계상금액"," 기부금명세서(A129)의 법정기부금(10), 정치자금(20), 문화진흥(60)의 합"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차가감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
				
				
				
		     Case "64"
		   	  '코드(64)50%손금기부금의 항목(3)회사계상금액	- 기부금명세서(A129)의 기부금(30)의 합과 일치 
		   	  '(코드(64)의 항목(3)의 금액이 “0”보다 큰 경우 A129 반드시 입력)
		
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
						
								Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W2, UNIGetMesg(lgcTB_47A.W2 & "_회사계상금액 '0'보다 큰 경우 기부금명세서(A129) 서식 필수 입력", "",""))
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '30' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "코드(64)50%손금기부금의 항목(3)회사계상금액","  기부금명세서(A129)의 기부금(30)의 합과 일치"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차가감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
				
				
		  Case "42"
		   	  '코드(42)지정기부금의 항목(3)회사계상금액	- 기부금명세서(A129)의 지정기부금(40), 문화단체(70)의 합과 일치 
		   	  '(코드(42)의 항목(3)의 금액이 “0”보다 큰 경우 A129 반드시 입력)
		   	  
		
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_회사계상금액 '0'보다 큰 경우 기부금명세서(A129) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '40' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							     cDataExists.A129.Find 2, "w9_cd = '70' " 
							     sTmp = sTmp + UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "코드(42)지정기부금의 항목(3)회사계상금액","  기부금명세서(A129)의 기타기부금의 항목(50) 합과 일치"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차가감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
				
				
				
			Case "73"
			   ' 코드(73)기타기부금의 항목(3)회사계상금액코드(73)기타기부금의 항목(4)세무상부인(조정)금액	
			    '- 기부금명세서(A129)의 마,기타기부금의 항목(50)과 일치 
			    '(코드(73)의 항목(3),코드(73)의 항목(4)의 금액이 “0”보다 큰 경우 A129 반드시 입력)
				
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
			        If UNICDbl(lgcTB_47A.W3,0) > 0 Then
			        	Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A129.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_회사계상금액 '0'보다 큰 경우 기부금명세서(A129) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
							
							    
							     cDataExists.A129.Find 2, "w9_cd = '50' " 
							     sTmp =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
							    
							    
								If UNICDbl(lgcTB_47A.W3, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "기타기부금의 항목(3)회사계상금액","  기부금명세서(A129)의 지정기부금(40), 문화단체(70)의 합과 일치"))
								End If
								
								If UNICDbl(lgcTB_47A.W4, 0) <> UNICDbl(sTmp , 0)  Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W4, UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_47A.W2 & "기타기부금의 항목(4)세무상부인(조정)금액","  기부금명세서(A129)의 지정기부금(40), 문화단체(70)의 합과 일치"))
								End If
								
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing				
					 End If		
			    
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
				
				
				
				
				
			Case "65"
			   '코드(65)접대비(5만원초과)의 항목(3)회사계상금액	
				'- 표준손익계산서 및 부속명세서의 접대비금액이 있는데 코드(65)의 항목(3) 의 금액이 없으면 오류 
				'(A115 코드(25), A116 코드(37), A117 코드(27), A118 코드(22), A119 코드(14), A120 코드(24), A121 코드(27), A123 코드(25))
				
			    If  ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
			      
			        	Set cDataExists.A115 = new C_TB_3_3	' -- W5109MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A115.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A115.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A115.LoadData() Then
								blnError = True
								Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "표준손익계산서 및 부속명세서의 접대비금액이 있는데 코드(65)의 항목(3) 의 금액이 없으면 오류 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
							
							    
								IF cDataExists.A115.W1 = 1 THEN 
								   cDataExists.A115.Find "w4 = 25 "
								ELSE
								   cDataExists.A115.Find "w4 = 37 "
								END IF      
								 
								sTmp =  UNICDbl(cDataExists.A115.W5 ,0)
									If UNICDbl(sTmp , 0) > 50000  and UNICDbl(lgcTB_47A.W3,0) = 0 Then
										blnError = True
										Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg("표준손익계산서의 접대비 항목에 금액이 있으면 코드(65)접대비(5만원초과)항목에 금액이 입력되어야합니다.","",""))
									End If
								
								
								
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A115 = Nothing				
			
			    Else
			       blnError = True	
			    End if   
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)	
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차가감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
				
				
				
			Case "66"
			
			    If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then blnError = True	
  		        sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)	
				
				
				
			   '코드(66)지정기부금한도액의 항목(5)차가감금액 
			   '- 법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준이 100만원보다 크고, 
			   '기부금명세서(A129)의 지정기부금_계가 100만원 이상일 경우 ‘0’보다 큰 값 입력 
				
			    If  ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "회사계상금액") Then 
						Set cDataExists.A101 = new C_TB_3	' -- W1109MA1_HTF.asp 에 정의됨 
									
							' -- 추가 조회조건을 읽어온다.
							'cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							'cDataExists.A101.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
									
							If Not cDataExists.A101.LoadData() Then
								blnError = True
								'Call SaveHTFError(lgsPGM_ID, "", lgcTB_47A.W2 & "_차가감금액 '0'보다 큰 경우 표준손익계산서 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							Else
								
								'sTmp = UNICDBL(cDataExists.A101.W10  ,0)
								sTmp = UNICDBL(cDataExists.A101.W56  ,0)	' -- 2006-01-05 : 200603 개정판 
							
							End If
					
								
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A101 = Nothing		
			    
			         if sTmp > 1000000 Then 
			    
							Set cDataExists.A129 = new C_TB_22	' -- W5109MA1_HTF.asp 에 정의됨 
											
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A129.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A129.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
							If Not cDataExists.A129.LoadData() Then
								'blnError = True
								'Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg( lgcTB_47A.W2 & "_회사계상금액 '0'보다 큰 경우 기부금명세서(A129) 서식 필수 입력 ","",""))		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
								sTmp = 0
							Else
									    
							     cDataExists.A129.Find 2, "w9_cd = '99' " 
							     sTmp2 =  UNICDbl(cDataExists.A129.GetData(2,"W9_AMT"),0)
										
							End If
								
										
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A129 = Nothing
					
							If sTmp2 > 1000000 and unicdbl(lgcTB_47A.W5,0)  <= 0 then
					
							     Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W3, UNIGetMesg( lgcTB_47A.W2 & "_차가감금액이 법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준이_ 100만원보다 크고 기부금명세서(A129)의 지정기부금_계가 100만원 이상일 경우 ‘0’보다 큰 값을 입력 " ,"",""))		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
							     blnError = True
							End If		
					End if					
			Else
			   	blnError = True
			End if		
			
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)	
				
			
				
			
			Case "49", "75", "76"
			
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)

			Case "47"
				If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
						
			Case Else
			
				If Not ChkNotNull(lgcTB_47A.W3, lgcTB_47A.W2 & "_" & "회사계상금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W3, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W4, lgcTB_47A.W2 & "_" & "세무조정금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W4, 15, 0)
						
				If Not ChkNotNull(lgcTB_47A.W5, lgcTB_47A.W2 & "_" & "차가감금액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W5, 15, 0)
		End Select
		
			
	
		'항목(5)차가감금액 =  항목(3)회사계상금액 - 항목(4)세무상부인(조정)금액(코드 53,12,71,13,72,55,56,57,58,59,41,64,42,65,61,74,77)
		SELECT CASE lgcTB_47A.W2_CD
		  CASE 53,12,71,13,72,55,56,57,58,59,41,64,42,65,61,74,77
				If UNICDbl(lgcTB_47A.W5,0)  <> UNICDbl(lgcTB_47A.W3,0) - UNICDbl(lgcTB_47A.W4,0) Then
				   Call SaveHTFError(lgsPGM_ID,lgcTB_47A.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL,lgcTB_47A.W2 &  "차가감금액","회사계상금액-세무조정금액"))
				   blnError = True	
				End if
		END SELECT  
		
		lgcTB_47A.MoveNext 
	Loop

	If  ChkNotNull(lgcTB_47A.W124, "상여배당등_소득처분금액") Then 
	
		'코드(97)소득처분금액	- 소득자료명세(A137)서 항목(4)소득금액_계와 일치(코드(97)이 “0”보다 큰 경우 A137 반드시 입력)
			
				    
		if UNICDbl(lgcTB_47A.W124,0) > 0  Then		    
				    
			Set cDataExists.A137 = new C_TB_55	' -- W5109MA1_HTF.asp 에 정의됨 
											
			' -- 추가 조회조건을 읽어온다.
			cDataExists.A137.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A137.WHERE_SQL = "and SEQ_NO = '999999' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
			If Not cDataExists.A137.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W124, "소득자료명세(A137) 서식이 작성되지 않았습니다.")
			Else
			
			     sTmp =  UNICDbl(cDataExists.A137.GetData("W4"),0)
							
											
			End If
								
										
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A137 = Nothing
						
			If  sTmp <> UNICDbl(lgcTB_47A.W124,0) then
			  	blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W124 & " <> " & sTmp, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"코드(97)소득처분금액","  소득자료명세(A137)서 항목(4)소득금액_계"))
			End If	
		End If					
	Else
	   	blnError = True
	End if		
			
	'200703
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '소비성서비스업영위법인_회사계상금액 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '소비성서비스업영위법인_세무조정금액 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '소비성서비스업영위법인_차가감금액 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '기타업 영위법인_회사계상금액 
	sHTFBody = sHTFBody & UNINumeric("0", 15, 0) '기타업 영위법인_차가감금액 
		
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W124, 15, 0)
						
	If Not ChkNotNull(lgcTB_47A.W125, "상여배당등_이익처분금") Then blnError = True	

	' -- 2006.03.24추가 : 법인구분 '2'는 검증제외 
	if UNICDbl(lgcTB_47A.W125,0) > 0 And lgcCompanyInfo.Comp_type2 <> "2" Then		    
				    
		Set cDataExists.A142 = new C_TB_3_3_4	' -- W5109MA1_HTF.asp 에 정의됨 
											
		' -- 추가 조회조건을 읽어온다.
		cDataExists.A142.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A142.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
		If Not cDataExists.A142.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W125, "이익잉여금처분(결손처리)계산서(A142) 서식이 작성되지 않았습니다.")
		Else
			
		     sTmp =  UNICDbl(cDataExists.A142.W5,0) + UNICDbl(cDataExists.A142.W15,0) + UNICDbl(cDataExists.A142.W26,0)
							
											
		End If
								
										
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A142 = Nothing
						
		If  sTmp <> UNICDbl(lgcTB_47A.W125,0) then
		  	blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_47A.W125 & " <> " & sTmp, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"코드(98)상여.배당등 이익처분금","  이익잉여금처분(결손처리)계산서(A142)의 코드(5)중간배당액 + 코드(15)배당금 + 코드(26)이익처분에의한상여금"))
		End If	
	End If					


	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47A.W125, 15, 0)
				
	sHTFBody = sHTFBody & UNIChar("", 49)	' -- 공란 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_47A = Nothing	' -- 메모리해제 
	
End Function


%>
