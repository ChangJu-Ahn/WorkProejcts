<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 별지국외특수관계자요약손익계산서 
'*  3. Program ID           : W9121MA1
'*  4. Program Name         : W9121MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_KJ_BJ1

Set lgcTB_KJ_BJ1 = Nothing	' -- 초기화 

Class C_TB_KJ_BJ1
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
			LoadData = False  
			Exit Function
		End If

		
		LoadData = True
	End Function

	Function EOF()
		EOF = lgoRs1.EOF
	End Function

	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	End Function	
		
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
			lgStrSQL = " SELECT  * FROM (" & vbCrLf
			
            lgStrSQL = lgStrSQL & " SELECT CONVERT(INT, A.W1 ) W1, A.W2, A.W3, A.W4, A.W5, A.W6  ,A.W7, A.W8, A.W9 ,  A.W10" & vbCrLf
            
			lgStrSQL = lgStrSQL & "  From  dbo.ufn_TB_KJ_BJ1_HOME_TAX_GetRef("& pCode1 &","& pCode2 &","& pCode3 &") A" & vbCrLf
			lgStrSQL = lgStrSQL & "  Union All " & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT  " & vbCrLf
            lgStrSQL = lgStrSQL & "  CONVERT(INT, A.W1 ) W1 ,  Cast(A.W2 as Varchar(15)), Cast(A.W3 as Varchar(15)), Cast(A.W4 as Varchar(15)), Cast(A.W5 as Varchar(15)), " & vbCrLf
            lgStrSQL = lgStrSQL & "   Cast(A.W6 as Varchar(15)),Cast(A.W7 as Varchar(15)), Cast(A.W8 as Varchar(15)), Cast(A.W9 as Varchar(15))  ,  Cast(A.W10 as Varchar(15)) " & vbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_KJ_BJ1 A WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & " ) X "	 & vbCrLf

			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

            lgStrSQL = lgStrSQL & " ORDER BY  X.W1 ASC" & vbcrlf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9121MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9121MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows, iSeqNo
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    Dim dblW01, dblW02, dblW03
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9121MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9121MA1"
	
	Set lgcTB_KJ_BJ1 = New C_TB_KJ_BJ1		' -- 해당서식 클래스 
	
	If Not lgcTB_KJ_BJ1.LoadData	Then Exit Function		' -- 제1호 서식 로드 
		
	'==========================================
	' -- 별지국외특수관계자요약손익계산서 및 오류검증 
	
	iSeqNo = 1

	For iDx = 2 To 10
		
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)	' 일련번호 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "명칭") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 40)
	
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "소재지_주소") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 70)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "사업연도(시작)") Then blnError = True		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_KJ_BJ1.GetData("W" & iDx))
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "사업연도(종료)") Then blnError = True		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_KJ_BJ1.GetData("W" & iDx))
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "주업종 코드") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 7)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "자본금액 또는 출자금액") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "특수관계의 구분") Then blnError = True		
		sHTFBody = sHTFBody & UNIChar(lgcTB_KJ_BJ1.GetData("W" & iDx), 1)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "주식등의 소유비율_소유 계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "주식등의 소유비율_소유직접") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "주식등의 소유비율_피소유 계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "주식등의 소유비율_피소유직접") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 6, 3)
		
		lgcTB_KJ_BJ1.MoveNext 
		lgcTB_KJ_BJ1.MoveNext ' -- 계정과목은 건너뜀: 2006.02.28 최영태수정(오류)
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "매출액") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW01 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "매출원가") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW02 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "매출총손익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		dblW03 = lgcTB_KJ_BJ1.GetData("W" & iDx)
		
		If dblW03 <> (dblW01 - dblW02) Then
			Call SaveHTFError(lgsPGM_ID, dblW03, UNIGetMesg(TYPE_CHK_NOT_EQUAL,"매출총손익", "매출액 - 매출원"))
		End If
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "판매비와관리비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "영업손익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		lgcTB_KJ_BJ1.MoveNext 
		
		If Not ChkNotNull(lgcTB_KJ_BJ1.GetData("W" & iDx), "법인세비용차감전손익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_KJ_BJ1.GetData("W" & iDx), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 25) & vbCrLf	' -- 공란 
	
		lgcTB_KJ_BJ1.MoveFirst 

		iSeqNo = iSeqNo + 1
		
		If iDx < 10 Then	' 마지막이 아닐때, 값이 없으면 루프탈출 
			If lgcTB_KJ_BJ1.GetData("W" & iDx+1) = "" Then Exit For
		End If
	Next
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_KJ_BJ1 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9121MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
	
	End Select
	PrintLog "SubMakeSQLStatements_W9121MA1 : " & lgStrSQL
End Sub

%>
