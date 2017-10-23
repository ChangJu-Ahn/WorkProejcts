<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 해외현지법인 명세서 
'*  3. Program ID           : W9127MA1
'*  4. Program Name         : W9127MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : leewolsan
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_A126

Set lgcTB_A126 = Nothing ' -- 초기화 

Class C_TB_A126
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
		Call GetData()
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
	            lgStrSQL = lgStrSQL & " A.*  , B.W7, B.W8" & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A126	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_A125	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO=B.SEQ_NO " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W01 > 0 AND ISNULL(B.W7,'') <>''"  & vbCrLf	' -- 데이타의 존재 유무 
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9127MA1
	Dim A126
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9127MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), sHTFBody2
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    
    PrintLog "MakeHTF_W9127MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9127MA1"

	Set lgcTB_A126 = New C_TB_A126		' -- 해당서식 클래스 
	
	If Not lgcTB_A126.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9127MA1

	' -- 쿼리변수 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

	'==========================================
	' -- 제4호 최저한세조정계산서 오류검증 
	iSeqNo = 1	: sHTFBody = ""

	Do Until lgcTB_A126.EOF 

		' -------------- 재무상황표 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & "A126"		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
        '3 일련번호 
		If Not ChkNotNull(lgcTB_A126.GetData("SEQ_NO"), "일련번호") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("SEQ_NO"), 6, 0)
			
	'4 현지법인명 

		If Not ChkNotNull(lgcTB_A126.GetData("W7"), "현지법인명") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A126.GetData("W7"), 60)

		'5 현지법인고유번호 
		If Not ChkNotNull(lgcTB_A126.GetData("W8"), "현지법인고유번호") Then blnError = True
		' -- 2006.03.29 개정  = 8 제외 
		' -- 첫글자가 1,2 가 아니면 
		If lgcTB_A126.GetData("W8") <> "99999999" And (Left(lgcTB_A126.GetData("W8"), 1) <> "1" And Left(lgcTB_A126.GetData("W8"), 1) <> "2"  And Left(lgcTB_A126.GetData("W8"), 1) <> "8") Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A126.GetData("W8"), UNIGetMesg("현지법인고유번호가 99999999가 아닐때, 첫글자가 1 또는 2 또는 8 이(가) 아니면 오류입니다", "",""))
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_A126.GetData("W8"), 8)
		
		'6.자산총계 
		
		If Not ChkNotNull(lgcTB_A126.GetData("W01"), "자산총계") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W01")), "자산총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W01"), 15, 0)
		
		
		'7.매출채권(특수관계기업)
	
		If Not ChkNotNull(lgcTB_A126.GetData("W03"), "매출채권(특수관계기업)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W03")), "매출채권(특수관계기업)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W03"), 15, 0)

		'8 매출채권(기타)
		If Not ChkNotNull(lgcTB_A126.GetData("W04"), "매출채권(기타)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W04")), "매출채권(기타)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W04"), 15, 0)
		
		'9 재고자산 
		If Not ChkNotNull(lgcTB_A126.GetData("W05"), "재고자산") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W05")), "재고자산") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W05"), 15, 0)
		
		'10 유가증권 
		If Not ChkNotNull(lgcTB_A126.GetData("W06"), "유가증권") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W06")), "유가증권") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W06"), 15, 0)
		
		
		'11 대여금(특수관계기업)
		If Not ChkNotNull(lgcTB_A126.GetData("W07"), "대여금(특수관계기업)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W07")), "대여금(특수관계기업)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W07"), 15, 0)
		
		'12 대여금(기타)
		If Not ChkNotNull(lgcTB_A126.GetData("W08"), "대여금(기타)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W08")), "대여금(기타)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W08"), 15, 0)
		
		'13 고정자산 
		If Not ChkNotNull(lgcTB_A126.GetData("W09"), "고정자산") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W09")), "고정자산") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W09"), 15, 0)
		
		'14토지및건축물 
		If Not ChkNotNull(lgcTB_A126.GetData("W10"), "토지및건축물") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W10")), "토지및건축물") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W10"), 15, 0)
		
		'15 기계장치,차량운반구 
		If Not ChkNotNull(lgcTB_A126.GetData("W11"), "기계장치,차량운반구") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W11")), "기계장치,차량운반구") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W11"), 15, 0)



		'16 고장자산 기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W12"), "고정자산 기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W12")), "고정자산 기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W12"), 15, 0)


		'17 무형자산 
		If Not ChkNotNull(lgcTB_A126.GetData("W13"), "무형자산") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W13")), "무형자산") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W13"), 15, 0)
		
		'18 자산기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W14"), "자산 기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W14")), "자산 기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W14"), 15, 0)
		
		'19 부채총계 
		If Not ChkNotNull(lgcTB_A126.GetData("W15"), "부채총계") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W15")), "부채총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W15"), 15, 0)
		
		'20 매입채무(특수관계기업)
		If Not ChkNotNull(lgcTB_A126.GetData("W16"), "매입채무(특수관계기업)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W16")), "매입채무(특수관계기업)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W16"), 15, 0)
		
		'21 매입채무(기타)
		If Not ChkNotNull(lgcTB_A126.GetData("W17"), "매입채무(기타)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W17")), "매입채무(기타)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W17"), 15, 0)


		'22 차입금(특수관계기업)
		If Not ChkNotNull(lgcTB_A126.GetData("W18"), "차입금(특수관계기업)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W18")), "차입금(특수관계기업)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W18"), 15, 0)


		'23 차입금(기타)
		If Not ChkNotNull(lgcTB_A126.GetData("W19"), "차입금(기타)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W19")), "자산총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W19"), 15, 0)
		
		'24 미지급금 
		If Not ChkNotNull(lgcTB_A126.GetData("W20"), "미지급금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W20")), "미지급금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W20"), 15, 0)
		
		'25 부채기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W21"), "부채 기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W21")), "부채 기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W21"), 15, 0)
		
		'26 자본금총계 
		If Not ChkNotNull(lgcTB_A126.GetData("W22"), "자본금총계") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W22")), "자본금총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W22"), 15, 0)
		
		'27 자본금 
		If Not ChkNotNull(lgcTB_A126.GetData("W23"), "자본금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W23")), "자본금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W23"), 15, 0)
		
		'28 기타자본금 
		If Not ChkNotNull(lgcTB_A126.GetData("W24"), "기타자본금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W24")), "기타자본금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W24"), 15, 0)
		
		'29 자본잉여금 
		If Not ChkNotNull(lgcTB_A126.GetData("W25"), "자본잉여금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W25")), "자본잉여금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W25"), 15, 0)
		
		'30 이익잉여금 
		If Not ChkNotNull(lgcTB_A126.GetData("W26"), "이익잉여금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W26")), "이익잉여금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W26"), 15, 0)
		
		'31 기타자본금 중 기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W27"), "기타자본금 중 기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W27")), "기타자본금 중 기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W27"), 15, 0)
		
		'32 매출액 
		If Not ChkNotNull(lgcTB_A126.GetData("W28"), "매출액") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W28")), "매출액") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W28"), 15, 0)
		
		'33 모기업 
		If Not ChkNotNull(lgcTB_A126.GetData("W29"), "매출액 모기업") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W29")), "매출액 모기업") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W29"), 15, 0)
		
		'34 매출액기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W30"), "매출액 기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W30")), "매출액 기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W30"), 15, 0)
		
		'35 매출원가 
		If Not ChkNotNull(lgcTB_A126.GetData("W31"), "매출원가") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W31")), "매출원가") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W31"), 15, 0)
		
		
		'36 판매비와관리비 
		If Not ChkNotNull(lgcTB_A126.GetData("W34"), "판매비와 일반관리비") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W34")), "판매비와 일반관리비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W34"), 15, 0)
		
		'37 급여(모회사파견직원)
		If Not ChkNotNull(lgcTB_A126.GetData("W35"), "급여(본사파견직원)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W35")), "급여(본사파견직원)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W35"), 15, 0)
		
		'38 급여(기타)
		If Not ChkNotNull(lgcTB_A126.GetData("W36"), "급여(기타)") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W36")), "급여(기타)") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W36"), 15, 0)
		
		
		'39 임차료 
		If Not ChkNotNull(lgcTB_A126.GetData("W37"), "임차료") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W37")), "임차료") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W37"), 15, 0)
		
		'40 연구개발비 
		If Not ChkNotNull(lgcTB_A126.GetData("W38"), "연구개발비") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W38")), "연구개발비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W38"), 15, 0)
		
		'41 대손상각비 
		If Not ChkNotNull(lgcTB_A126.GetData("W39"), "대손상각비") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W39")), "대손상각비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W39"), 15, 0)
		
		'42 판매비와 일반관리비-기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W40"), "판매비와 일반관리비_기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W40")), "판매비와 일반관리비_기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W40"), 15, 0)
		
		'43 영업외수익 
		If Not ChkNotNull(lgcTB_A126.GetData("W41"), "영업외수익") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W41")), "영업외수익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W41"), 15, 0)
		
		'44 이자수익 
		If Not ChkNotNull(lgcTB_A126.GetData("W42"), "이자수익") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W42")), "이자수익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W42"), 15, 0)
		
		'45 배당수익 
		If Not ChkNotNull(lgcTB_A126.GetData("W43"), "배당수익") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W43")), "배당수익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W43"), 15, 0)
		
		'46 영업외수익-기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W44"), "영업외수익_기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W44")), "영업외수익_기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W44"), 15, 0)
		
		
		'47 영업외비용 
		If Not ChkNotNull(lgcTB_A126.GetData("W45"), "영업외비용") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W45")), "영업외비용") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W45"), 15, 0)
		
		
		'48 이자비용 
		If Not ChkNotNull(lgcTB_A126.GetData("W46"), "이자비용") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W46")), "이자비용") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W46"), 15, 0)
		
		
		'49 영업외비용-기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W47"), "영업외비용_기타") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W47")), "영업외비용_기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W47"), 15, 0)
		
		'50 특별이익 
		If Not ChkNotNull(lgcTB_A126.GetData("W48"), "특별이익") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W48")), "특별이익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W48"), 15, 0)
		
		'51 특별손실 
		If Not ChkNotNull(lgcTB_A126.GetData("W51"), "특별손실") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W51")), "특별손실") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W51"), 15, 0)
		
		'52 법인세 
		If Not ChkNotNull(lgcTB_A126.GetData("W52"), "법인세") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W52")), "법인세") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W52"), 15, 0)
		
		'53 당기순손익 
		If Not ChkNotNull(lgcTB_A126.GetData("W53"), "당기순손익") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W53")), "당기순손익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W53"), 15, 0)
		
		'54 현금과예금 
		If Not ChkNotNull(lgcTB_A126.GetData("W02"), "현금과예금") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W02")), "현금과예금") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W02"), 15, 0)
		
		
		'55 매출원가-모기업으로부터매입 
		If Not ChkNotNull(lgcTB_A126.GetData("W32"), "매출원가-모기업으로부터매입") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W32")), "매출원가-모기업으로부터매입") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W32"), 15, 0)
		
		
		'56 매출원가- 기타매입 
		If Not ChkNotNull(lgcTB_A126.GetData("W33"), "매출원가- 기타매입") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W33")), "매출원가- 기타매입") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W33"), 15, 0)
		
		'57 특별이익 - 채무면제익 
		If Not ChkNotNull(lgcTB_A126.GetData("W49"), "매출원가- 기타매입") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W49")), "매출원가- 기타매입") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W49"), 15, 0)
		
		'57 특별이익 - 기타 
		If Not ChkNotNull(lgcTB_A126.GetData("W50"), "매출원가- 기타매입") Then blnError = True
		If Not ChkNumeric(CStr(lgcTB_A126.GetData("W50")), "매출원가- 기타매입") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A126.GetData("W50"), 15, 0)

		'58 공란 25
		sHTFBody = sHTFBody & UNIChar("", 25) 
		
		If Not blnError Then
			Call WriteLine2File(sHTFBody)
		End If
		sHTFBody=""
		lgcTB_A126.MoveNext 
	Loop


	PrintLog "WriteLine2File : " & sHTFBody
	
	' -- 파일에 기록한다.

	If Not blnError Then

		'Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_A126 = Nothing	' -- 메모리해제 

End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9127MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A126" '-- 외부 참조 SQL

			lgStrSQL = ""
			

	End Select
	PrintLog "SubMakeSQLStatements_W9127MA1 : " & lgStrSQL
End Sub
%>
