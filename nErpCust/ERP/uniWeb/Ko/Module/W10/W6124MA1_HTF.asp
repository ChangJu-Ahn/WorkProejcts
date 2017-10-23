<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제8호 공제감면세액계산서(1)
'*  3. Program ID           : W6124MA1
'*  4. Program Name         : W6124MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_8_1

Set lgcTB_8_1 = Nothing	' -- 초기화 

Class C_TB_8_1
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
				lgStrSQL = lgStrSQL & " FROM TB_8_1	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6124MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6124MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6124MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6124MA1"
	
	Set lgcTB_8_1 = New C_TB_8_1		' -- 해당서식 클래스 
	
	If Not lgcTB_8_1.LoadData	Then Exit Function		' -- 제1호 서식 로드 
		
	'==========================================
	' -- 제8호 공제감면세액계산서(1) 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_1"), "공공차관도입에 따른 법인세감면_공제감면세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_1"), 15, 0)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_2"), "재해손실세액공제_공제감면세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_2"), 15, 0)
	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W1_3"), 40)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W4_3"), "기타1_공제감면세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_3"), 15, 0)
	
	If  ChkNotNull(lgcTB_8_1.GetData("W4_SUM"), "공제감면세액_계") Then 
		if unicdbl(lgcTB_8_1.GetData("W4_SUM"),0) <> unicdbl(lgcTB_8_1.GetData("W4_1"),0) + unicdbl(lgcTB_8_1.GetData("W4_2"),0) + unicdbl(lgcTB_8_1.GetData("W4_3"),0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2.GetData("W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, " 공제감면세액_계","공제감면세액의 총합"))
		      blnError = True		
		End if
	Else
	   blnError = True		
	End if
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W4_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_1"), "재해내용") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W5_1"), 40)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_2"), "재해발생일") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_8_1.GetData("W5_2"))

	If Not ChkNotNull(lgcTB_8_1.GetData("W5_3"), "공제신청일") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_8_1.GetData("W5_3"))

	If Not ChkNotNull(lgcTB_8_1.GetData("W5_4_GB"), "구분") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_8_1.GetData("W5_4_GB"), 20)
	
	If Not ChkNotNull(lgcTB_8_1.GetData("W5_4"), "법인세") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_8_1.GetData("W5_4"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 53)	' -- 공란 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_8_1 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6124MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
	
	End Select
	PrintLog "SubMakeSQLStatements_W6124MA1 : " & lgStrSQL
End Sub

%>
