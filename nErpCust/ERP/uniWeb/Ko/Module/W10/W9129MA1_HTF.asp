
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 해외현지지사 명세서 
'*  3. Program ID           : W9129MA1
'*  4. Program Name         : W9129MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : lee wol san
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_A128

Set lgcTB_A128 = Nothing ' -- 초기화 

Class C_TB_A128
	' -- 테이블의 컬럼변수 
	
	Dim WHERE_SQL		' -- 기본 검색조건(지사/사업연도/신고구분)외의 검색조건 
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
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_A128	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W6 <> '' "  & vbCrLf	' -- 데이타의 존재 유무 
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9129MA1
	Dim A126
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9129MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), blnChkA126A127
    DIM oRs3,oRs4
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    
    PrintLog "MakeHTF_W9129MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9129MA1"

	Set lgcTB_A128 = New C_TB_A128		' -- 해당서식 클래스 
	
	If Not lgcTB_A128.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9129MA1

	' -- 쿼리변수 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

	'==========================================
	' -- 제4호 최저한세조정계산서 오류검증 
	iSeqNo = 1	: sHTFBody = ""

	Do Until lgcTB_A128.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 


		'3 페이지번호 
		If Not ChkNotNull(lgcTB_A128.GetData("SEQ_NO"), "페이지번호") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("SEQ_NO"), 4, 0)
		
		'4 전기말가동지사수 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W15"), 5, 0)
		
		'5 당기신설지사수 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W16"), 5, 0)
		
		'6 당기폐쇄지사수 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W17"), 5, 0)
		
		'7 당기말가동지사수 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W18"), 5, 0)
		
		'8 소재지국 
		If Not ChkNotNull(lgcTB_A128.GetData("W6"), "국카코드") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W6"), 3)
		
		'9 해외지사명 
		If Not ChkNotNull(lgcTB_A128.GetData("W7"), "현지지사명") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W7"), 60)
		
		'10 해외지사고유번호 
		If Not ChkNotNull(lgcTB_A128.GetData("W8"), "현지지사고유번호") Then blnError = True	
		If Len(lgcTB_A128.GetData("W8")) <> 8 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W8"), UNIGetMesg("전체길이가 8이 아니면 오류입니다.", "",""))
		End If
		
		If Left(lgcTB_A128.GetData("W8"), 1) <> "9" And Left(lgcTB_A128.GetData("W8"), 1) <> "8" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W8"), UNIGetMesg("첫글자가 9, 8 이(가) 아니면 오류입니다", "",""))
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W8"), 8)
		
		
		'11 실립형태 
		If Not ChkBoundary("1,2", lgcTB_A128.GetData("W9"), "설립형태: " & lgcTB_A128.GetData("W9") & " " ) Then blnError = True
      
		Call SubMakeSQLStatements_W9129MA1("A",lgcTB_A128.GetData("SEQ_NO"),iKey1,iKey2, iKey3)  '
        If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("해외지사 경영현황을 입력하세요.", "",""))
		End If
		if lgcTB_A128.GetData("W9")="1" then '지점 

			if oRs3("allSum")="0" then '모든 sum이 0 이면 오류 
				Call SaveHTFError(lgsPGM_ID, oRs3("allSum"), UNIGetMesg("지점인 경우 자산총계~당기순이익의 금액이 0이면 오류입니다. ", "",""))
			end if
		elseif  lgcTB_A128.GetData("W9")="2" then '사무소 
			if oRs3("allSum")<>"0" then '모든 sum이 0 <> 이면 오류 
				Call SaveHTFError(lgsPGM_ID, oRs3("allSum"), UNIGetMesg("사무소인 경우 자산총계~당기순이익의 금액이 0이어야 합니다. ", "",""))
			end if
		else
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W9"), UNIGetMesg("1 또는 2 이어야 합니다. ", "",""))
		end if
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W9"), 1)
		
		'12 설립일자 
		If Not ChkNotNull(lgcTB_A128.GetData("W10"), "설립일자") Then blnError = True
		If DateDiff("m", lgcTB_A128.GetData("W10"), lgcTB_A128.GetData("W21")) < 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W10"), UNIGetMesg("설립일자는 폐쇄일보다 작아야합니다.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A128.GetData("W10"))
		
		'13 해외지사소재지 
		If Not ChkNotNull(lgcTB_A128.GetData("W11"), "현지지사 소재지") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W11"), 70)
		
		'14 업종코드 
		 Call SubMakeSQLStatements_W9129MA1("2",lgcTB_A128.GetData("W12"), "", "")  '업종코드 
		
		If   FncOpenRs("R",lgObjConn, oRs4, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("업종코드가 존재하지 않습니다.", "",""))
		End If
			Call SubCloseRs(oRs4)
		If Not ChkNotNull(lgcTB_A128.GetData("W12"), "업종코드") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A128.GetData("W12"), 7)


		'15 직원수 
		If Not ChkNotNull(lgcTB_A128.GetData("W13"), "직원수") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W13")), "직원수") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W13"), 5, 0)
		
		'16 본점파견직원수 
		
		If Not ChkNotNull(lgcTB_A128.GetData("W14"), "본점파견직원수") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W14")), "본점파견직원수") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W14"), 5, 0)
		
		


		
		'=========================================================================
		
		'17 자산총계 
		If Not ChkNotNull(oRs3("W1"), "자산총계") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W1")), "자산총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W1"), 15, 0)

		'18 토지및건축물 
		If Not ChkNotNull(oRs3("W2"), "토지및건축물") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W2")), "토지및건축물") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W2"), 15, 0)

		'19 기계장치,차랑운반구 
		If Not ChkNotNull(oRs3("W3"), "기계장치,차랑운반구") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W3")), "기계장치,차랑운반구") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W3"), 15, 0)
	
		'20 자산기타 
		If Not ChkNotNull(oRs3("W4"), "자산기타") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W4")), "자산기타") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W4"), 15, 0)

		'21 부채총계 
		If Not ChkNotNull(oRs3("W5"), "부채총계") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W5")), "부채총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W5"), 15, 0)

		'22 자본총계 
		If Not ChkNotNull(oRs3("W6"), "자본총계") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W6")), "자본총계") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W6"), 15, 0)


		'23 본점지원경비 
		If Not ChkNotNull(oRs3("W7"), "본점지원경비") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W7")), "본점지원경비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W7"), 15, 0)


		'24 매출액 
		If Not ChkNotNull(oRs3("W8"), "매출액") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W8")), "매출액") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W8"), 15, 0)

		'25 매출원가 
		If Not ChkNotNull(oRs3("W9"), "매출원가") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W9")), "매출원가") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W9"), 15, 0)

		'26 판매비와일반과리비 
		If Not ChkNotNull(oRs3("W10"), "판매비와일반과리비") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W10")), "판매비와일반과리비") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W10"), 15, 0)

		'27 영업외수익 
		If Not ChkNotNull(oRs3("W11"), "영업외수익") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W11")), "영업외수익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W11"), 15, 0)

		'28 영업외비용 
		If Not ChkNotNull(oRs3("W12"), "영업외비용") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W12")), "영업외비용") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W12"), 15, 0)

		'29 특별이익 
		If Not ChkNotNull(oRs3("W13"), "특별이익") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W13")), "특별이익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W13"), 15, 0)

		'30 특별손실 
		If Not ChkNotNull(oRs3("W14"), "특별손실") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W14")), "특별손실") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W14"), 15, 0)

		'31 법인세 
		If Not ChkNotNull(oRs3("W15"), "법인세") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W15")), "법인세") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W15"), 15, 0)

		'32 당기순손익 
		If Not ChkNotNull(oRs3("W16"), "당기순손익") Then blnError = True	
		If Not ChkNumeric(CStr(oRs3("W16")), "당기순손익") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(oRs3("W16"), 15, 0)
		Call SubCloseRs(oRs3)
		'=========================================================================
		'33 폐쇄일 
		If CDbl(lgcTB_A128.GetData("W22")) > 0 Then
			If Not ChkNotNull(lgcTB_A128.GetData("W21"), "폐쇄일") Then 
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W21"), UNIGetMesg("회수금액이 0보다 크면 폐쇄일은 반드시 입력되어야 합니다.", "",""))
			End If
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A128.GetData("W21"))

		'34 회수금액원화 
		
		If Not ChkNotNull(lgcTB_A128.GetData("W22"), "회수금액원화") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A128.GetData("W22")), "회수금액원화") Then blnError = True		
		If IsDate(lgcTB_A128.GetData("W21")) Then
			If CDbl(lgcTB_A128.GetData("W22")) <= 0 Then
				blnError = True	
				Call SaveHTFError(lgsPGM_ID, lgcTB_A128.GetData("W22"), UNIGetMesg("폐쇄일이 입력되면 회수금액이 0보다 커야 합니다.", "",""))
			End If
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A128.GetData("W22"), 15, 0)
		'35 공란 
		sHTFBody = sHTFBody & UNIChar("", 40) ' -- 공란	 :
		If Not blnError Then
			Call WriteLine2File(sHTFBody)
		End If
		sHTFBody=""
		
		lgcTB_A128.MoveNext 
	Loop

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.

	
	
	If Not blnError Then
		'Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_A128 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9129MA1(pMode,pVal, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A126" '-- 외부 참조 SQL

			lgStrSQL = ""
			
      Case "A"
      
		lgStrSQL =  " select max(w1) w1,max(w2 ) w2 ,max(w3 ) w3 ,max(w4 ) w4 ,max(w5 ) w5 ,max(w6 ) w6 ," & CHR(13)
		lgStrSQL = lgStrSQL & " max(w7 ) w7 ,max(w8 ) w8 ,max(w9 ) w9 ,max(w10) w10,max(w11) w11,max(w12) w12,    " & CHR(13)
		lgStrSQL = lgStrSQL & " max(w13) w13,max(w14) w14,max(w15) w15,max(w16) w16, sum(allSum) allSum                               " & CHR(13)
		lgStrSQL = lgStrSQL & " from (                                                                            " & CHR(13)
		lgStrSQL = lgStrSQL & " select  case when w2='01' then max(w3) end w1,                                    " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='02' then max(w3) end w2 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='03' then max(w3) end w3 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='04' then max(w3) end w4 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='05' then max(w3) end w5 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='06' then max(w3) end w6 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='07' then max(w3) end w7 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='08' then max(w3) end w8 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='09' then max(w3) end w9 ,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='10' then max(w3) end w10,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='11' then max(w3) end w11,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='12' then max(w3) end w12,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='13' then max(w3) end w13,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='14' then max(w3) end w14,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='15' then max(w3) end w15,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  case when w2='16' then max(w3) end w16,                                          " & CHR(13)
		lgStrSQL = lgStrSQL & "  sum(w3) allSum																	  " & CHR(13)

		lgStrSQL = lgStrSQL & " from  tb_a128_1                                                                   " & CHR(13)
		lgStrSQL = lgStrSQL & " where  tb_a128_1.co_cd=" &  pCode1   & "                                          " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.fisc_year=" &  pCode2   & "                                         " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.rep_type=" &  pCode3   & "                                                    " & CHR(13)
		lgStrSQL = lgStrSQL & " and tb_a128_1.seq_no = '"&pVal&"'                                                        " & CHR(13)
		lgStrSQL = lgStrSQL & " group by w2 ) a                                                                   " & CHR(13)

	Case "2" '업종체크 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " select top 1 STD_INCM_RT_CD from tb_std_income_rate"  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE ATTRIBUTE_YEAR = '2005' " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND STD_INCM_RT_CD= " & filterVar(pCode1 ,"''","S")	 & vbCrLf
			
	End Select
	PrintLog "SubMakeSQLStatements_W9129MA1 : " & lgStrSQL
End Sub
%>
