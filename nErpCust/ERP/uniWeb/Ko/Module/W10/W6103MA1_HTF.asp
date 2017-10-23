<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 조특 제1호 세액공제신청서 
'*  3. Program ID           : W6103MA1
'*  4. Program Name         : W6103MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_JT1

Set lgcTB_JT1 = Nothing ' -- 초기화 

Class C_TB_JT1
	' -- 테이블의 컬럼변수 
	
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
		lgoRs1.MoveFirst
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
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_JT1	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6103MA1
	Dim A106

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6103MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum,st1_SumTax,st1_Tax
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6103MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6103MA1"

	Set lgcTB_JT1 = New C_TB_JT1		' -- 해당서식 클래스 
	
	If Not lgcTB_JT1.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 


	'==========================================
	' -- 제15호 소득금액조정합계표 오류검증 
	iSeqNo = 1	: sHTFBody = ""
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
'	Do Until lgcTB_JT1.EOF 
	
'		Select Case lgcTB_JT1.GetData("W3")
'			Case "78"	' -- 기타 
'				sHTFBody = sHTFBody & UNIChar(lgcTB_JT1.GetData("W1"), 30)	' 널허용 
'				
'				If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
'				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
'				
'				If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
'				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
'				st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)
'				
'			Case "50"	' -- 합계 
'				If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
'				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
'				st1_Tax = unicdbl(lgcTB_JT1.GetData("W6"),0)
'				
'			Case Else
'				If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
'				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
'				
'				If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
'				sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
'				st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)
'		End Select
'		lgcTB_JT1.MoveNext 
'	Loop
		
	lgcTB_JT1.FIND " W3='31' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='93' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='75' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='78' "	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT1.GetData("W1"), 30)	' 널허용 
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='32' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='85' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='34' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='76' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='35' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='36' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='77' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='37' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='91' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='42' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='84' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='94' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='95' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='96' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W5"), lgcTB_JT1.GetData("W1") & "_대상세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W5"), 15, 0)
				
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_SumTax = st1_SumTax + unicdbl(lgcTB_JT1.GetData("W6"),0)

	lgcTB_JT1.FIND " W3='50' "	
	If Not ChkNotNull(lgcTB_JT1.GetData("W6"), lgcTB_JT1.GetData("W1") & "_공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT1.GetData("W6"), 15, 0)
	st1_Tax = unicdbl(lgcTB_JT1.GetData("W6"),0)

	
	sHTFBody = sHTFBody & UNIChar("", 9) & vbCrLf	' -- 공란 
	
	if unicdbl(st1_SumTax,0) <> unicdbl(st1_Tax,0)  then
	     Call SaveHTFError(lgsPGM_ID, st1_SumTax, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "항목들의 감면세액합","(118)감면세액"))
	     blnError = True		
	end if 
	' ----------- 
	'Call SubCloseRs(oRs2) 
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	

	Set lgcTB_JT1 = Nothing	' -- 메모리해제 
	
End Function





' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6103MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A106_93" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '93' " 	 & vbCrLf
           
      Case "A106_75" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '75' " 	 & vbCrLf	     
      
     Case "A106_32" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '32' " 	 & vbCrLf	  
     
      Case "A106_91" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '91' " 	 & vbCrLf	                       	
      
      Case "A106_94" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '94' " 	 & vbCrLf
          
      Case "A106_95" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '95' " 	 & vbCrLf	            	      
       
       Case "A106_96" '-- 외부 참조 SQL
			lgStrSQL = ""
            lgStrSQL = lgStrSQL & "	AND A.W2_1	= '96' " 	 & vbCrLf	           
            
	  
			
	End Select
	PrintLog "SubMakeSQLStatements_W6103MA1 : " & lgStrSQL
End Sub
%>
