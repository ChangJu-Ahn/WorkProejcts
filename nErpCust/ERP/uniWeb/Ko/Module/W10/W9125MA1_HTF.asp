
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 해외현지법인 명세서 
'*  3. Program ID           : W9125MA1
'*  4. Program Name         : W9125MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2006/01/19
'*  7. Modified date(Last)  : 2007/03
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : LEEWOLSAN
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_A125
Dim lgcTB_A125_1
Dim lgcTB_A125_2


Set lgcTB_A125 = Nothing ' -- 초기화 
Set lgcTB_A125_1 = Nothing ' -- 초기화 
Set lgcTB_A125_2 = Nothing ' -- 초기화 
Set lgcCompanyInfo = Nothing ' -- 초기화 




'===========================================================================
'C_TB_A125
'===========================================================================
  
  
Class C_TB_A125
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
			GetData		= lgoRs1(pFieldNm).value
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
				lgStrSQL = lgStrSQL & " FROM TB_A125	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
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
Class TYPE_DATA_EXIST_W9125MA1
	Dim A126
End Class


'===========================================================================
'C_TB_A125_1
'===========================================================================
  
Class C_TB_A125_1
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
				Call SaveHTFError(lgsPGM_ID, "_투자현황", TYPE_DATA_NOT_FOUND)
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
			GetData		= lgoRs1(pFieldNm).value
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
				lgStrSQL = lgStrSQL & " FROM TB_A125_1	A  WITH (NOLOCK) JOIN TB_A125 B  " & vbCrLf	' 
				lgStrSQL = lgStrSQL & " ON A.CO_CD=B.CO_CD AND A.FISC_YEAR= B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO = B.SEQ_NO" 	 & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W2 + A.W3 + A.W4+A.W5+A.W6 <>0 " 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		ORDER BY A.SEQ_NO,SEQ " 	 & vbCrLf



				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class


'===========================================================================
'C_TB_A125_2
'===========================================================================
  
Class C_TB_A125_2
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
				'Call SaveHTFError(lgsPGM_ID, "_자회사현황", TYPE_DATA_NOT_FOUND)
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
			GetData		= lgoRs1(pFieldNm).value
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
				lgStrSQL = lgStrSQL & " FROM TB_A125_2	A  WITH (NOLOCK) JOIN TB_A125 B  " & vbCrLf	' 
				lgStrSQL = lgStrSQL & " ON A.CO_CD=B.CO_CD AND A.FISC_YEAR= B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.SEQ_NO = B.SEQ_NO" 	 & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.W4+A.W5+A.W6 <>0 " 	 & vbCrLf
				'lgStrSQL = lgStrSQL & "		AND ISNULL(W1,'')<>''" 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		ORDER BY A.SEQ_NO,SEQ " 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class


' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9125MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25), blnChkA126A127
    Dim oRs3
    
'    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False	: blnChkA126A127 = False
    PrintLog "MakeHTF_W9125MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9125MA1"

	Set lgcTB_A125 = New C_TB_A125		' -- 해당서식 클래스 
	Set lgcTB_A125_1 = New C_TB_A125_1		' -- 해당서식 클래스 
	Set lgcTB_A125_2 = New C_TB_A125_2		' -- 해당서식 클래스 
	
	If Not lgcTB_A125.LoadData Then Exit Function			

	Call lgcTB_A125_1.LoadData
	Call lgcTB_A125_2.LoadData

	'If Not lgcCompanyInfo.LoadData Then Exit Function			' -- 법인기초정보 로드 
	
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9125MA1

	' -- 쿼리변수 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

	'==========================================

	iSeqNo = 1	: sHTFBody = ""
	
	'==========================================
	'해외현집법인 - 기본사항 
	'==========================================
	dim tmpVal
	Do Until lgcTB_A125.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
        '3. 페이지번호 
        If Not ChkNotNull(lgcTB_A125.GetData("SEQ_NO"), "페이지번호") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("SEQ_NO"), 4, 0)
		
        '4 전기말가동법인수 
        If Not ChkNumeric(CStr(lgcTB_A125.GetData("W15")), "전기말가동법인수") Then blnError = True
        sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W15"), 5, 0)
        
        '5 당기신설법인수 
         If Not ChkNumeric(CStr(lgcTB_A125.GetData("W16")), "당기신설법인수") Then blnError = True
        sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W16"), 5, 0)
   
        '6 당기청산법인수 
         If Not ChkNumeric(CStr(lgcTB_A125.GetData("W17")), "당기신설법인수") Then blnError = True
         sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W17"), 5, 0)

        '7 당기말가동법인수 
        tmpVal= CDbl(lgcTB_A125.GetData("W15")) + cdbl(lgcTB_A125.GetData("W16")) - cDbl(lgcTB_A125.GetData("W17"))
        
         If Not ChkNumeric(tmpVal, "당기말가동법인수") Then blnError = True

         sHTFBody = sHTFBody & UNINumeric(tmpVal, 5, 0)

        '8 재무상황표제출법인수 
        If Not ChkNumeric(CStr(lgcTB_A125.GetData("W18")), "재무상황표제출법인수") Then blnError = True
         sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W17"), 5, 0)

        '9 투자국코드 
        If Not ChkNotNull(lgcTB_A125.GetData("W6"), "투자국") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W6"), 3)
		
        '10 현지법인명 
        If Not ChkNotNull(lgcTB_A125.GetData("W7"), "현지법인명") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W7"), 60)
		
        '11 현지법인고유번호 
        If Not ChkNotNull(lgcTB_A125.GetData("W8"), "현지법인고유번호") Then blnError = True
		If Len(lgcTB_A125.GetData("W8")) <> 8 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W8"), UNIGetMesg("전체길이가 8이 아니면 오류입니다.", "",""))
		End If
			' -- 2006.03.29 개정  = 8 제외 
			' -- 첫글자가 1,2 가 아니면 
		if lgcTB_A125.GetData("W8") <> "99999999" And (Left(lgcTB_A125.GetData("W8"), 1) <> "1" And Left(lgcTB_A125.GetData("W8"), 1) <> "2" And Left(lgcTB_A125.GetData("W8"), 1) <> "8" ) Then
		
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W8"), UNIGetMesg("현지법인고유번호가 99999999가 아닐때, 첫글자가 1 또는 2 또는 8 이(가) 아니면 오류입니다", "",""))
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W8"), 8)
		
		
        '12 현지법인소재지 
        
        If Not ChkNotNull(Replace(lgcTB_A125.GetData("W9"),vbCrLf,""), "현지법인소재지") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(Replace(lgcTB_A125.GetData("W9"),vbCrLf,""), 70)
		
        '13 설립일자 
        If Not ChkNotNull(lgcTB_A125.GetData("W10"), "설립일자") Then blnError = True
		If DateDiff("m", lgcTB_A125.GetData("W10"), lgcTB_A125.GetData("W11_1")) < 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W10"), UNIGetMesg("설립일자는 사업연도_시작일보다 같거나 작아야합니다.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W10"))
		
       ' 14 사업연도_시작일 
       If Not ChkNotNull(lgcTB_A125.GetData("W11_1"), "사업연도_시작일") Then blnError = True
		If DateDiff("m", lgcTB_A125.GetData("W11_1"), lgcTB_A125.GetData("W11_2")) <= 0 Then 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W11_1"), UNIGetMesg("사업연도_시작일은 사업연도_종료일보다 작아야합니다.", "",""))
		End If
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W11_1"))
		
		
        '15 사업연도_종료일 
        If Not ChkNotNull(lgcTB_A125.GetData("W11_1"), "사업연도_종료일") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W11_2"))
		
        '16 업종코드 
        Call SubMakeSQLStatements_W9125MA1("2",lgcTB_A125.GetData("W12"), "", "")  '업종체크 
        If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("업종코드가 존재하지 않습니다.", "",""))
		End If

        
        If Not ChkNotNull(lgcTB_A125.GetData("W12"), "업종코드") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_A125.GetData("W12"), 7)
		
		
        '17 직원수 
        If Not ChkNotNull(lgcTB_A125.GetData("W13"), "직원수") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A125.GetData("W13")), "직원수") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W13"), 5, 0)
		
        '18 모법인파견직원수 
        
		If Not ChkNotNull(lgcTB_A125.GetData("W14"), "모법인파견직원수") Then blnError = True	
		If Not ChkNumeric(CStr(lgcTB_A125.GetData("W14")), "모법인파견직원수") Then blnError = True		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W14"), 5, 0)
		
        '19 청산-청산일 
        sHTFBody = sHTFBody & UNI8Date(lgcTB_A125.GetData("W19"))	' -- Null 허용 
        
        '20청산- 회수금액 
        If CDbl(lgcTB_A125.GetData("W20")) > 0 And lgcTB_A125.GetData("W19") = "" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W16"), UNIGetMesg("회수금액이 0보다 크면 청산일이 기재되어 있어야 합니다.", "",""))
		End If
		
        If IsDate(lgcTB_A125.GetData("W19")) And CDbl(lgcTB_A125.GetData("W20")) <= 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_A125.GetData("W20"), UNIGetMesg("청산일자 입력시 회수금액은 0보다 커야 됩니다.", "",""))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125.GetData("W20"), 15, 0)
		
        '21 공란 60
		sHTFBody = sHTFBody & UNIChar("", 60)  & vbCrLf ' -- 공란	 :

		lgcTB_A125.MoveNext 
	Loop


	'==========================================
	'2 해외현집법인 - 투자현황 
	'==========================================
	
	
	Do Until lgcTB_A125_1.EOF 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
        
        '3. 페이지번호 
         If Not ChkNotNull(lgcTB_A125_1.GetData("SEQ_NO"), "페이지번호") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("SEQ_NO"), 4, 0)
          
	
        '4.일련번호 
         If Not ChkNotNull(lgcTB_A125_1.GetData("SEQ"), "일련번호") Then blnError = True
        
			'--대부수입이자체크 
			IF lgcTB_A125_1.GetData("SEQ")<>"1" then

				if lgcTB_A125_1.GetData("W5")<>"0" or lgcTB_A125_1.GetData("W6")<>"0" then

					Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W5") & ":" & lgcTB_A125_1.GetData("w6"), UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "대여금,대부수입이자",""))
					blnError = True
		      	end if
			end if
			
			IF lgcTB_A125_1.GetData("SEQ")="999999" then
				'Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W1"), UNIGetMesg(TYPE_CHK_ZERO_EQUAL, "대여금,대부수입이자",""))
				'blnError = True
			end if
			
		sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("SEQ"), 6, 0)
        
        '5.(현지법인)주주명 NULL 허용 

        IF lgcTB_A125_1.GetData("SEQ")="1" then '000001인경우 A100법인명과 일치해야함. '
       
			if lgcCompanyInfo.CO_NM<>lgcTB_A125_1.GetData("W1") then
				Call SaveHTFError(lgsPGM_ID, lgcCompanyInfo.CO_NM & ":" & lgcTB_A125_1.GetData("w1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "법인명","(현지법인)주주명"))
			end if
        ELSE
        END IF 
         'If Not ChkNotNull(lgcTB_A125_1.GetData("W1"), "(현지법인)주주명") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_1.GetData("W1"), 30)
		
        '6.출자금액 
        
        If Not ChkNotNull(lgcTB_A125_1.GetData("W2"), "출자금액") Then blnError = True	
		  sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W2"), 15, 0)
		 
        '7출자비율 
        
			IF lgcTB_A125_1.GetData("SEQ")>"1" and lgcTB_A125_1.GetData("SEQ") < "999998"   then
				IF lgcTB_A125_1.GetData("W3")<"10" then
				Call SaveHTFError(lgsPGM_ID, lgcTB_A125_1.GetData("W3"), UNIGetMesg(TYPE_CHK_OVER_EQUAL, "출자비율","10%"))
				end if
			end if
        
         If Not ChkNotNull(lgcTB_A125_1.GetData("W3"), "출자비율") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W3"), 5, 1)
		 
        '8 배당금수입 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W3"), "배당금수입") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W3"), 15, 0)
		 
        '9 대여금 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W5"), "대여금") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W5"), 15, 0)
 
        '10 대부수입이자 
         If Not ChkNotNull(lgcTB_A125_1.GetData("W6"), "대부수입이자") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_1.GetData("W6"), 15, 0)
		
        '11공란 39
         sHTFBody = sHTFBody & UNIChar("", 39) & vbCrLf' -- 공란	 :
		
		lgcTB_A125_1.MoveNext 
		
	Loop
	

	
	'==========================================
	'3 해외현집법인 - 자회사현황 
	'==========================================
	
	Do Until lgcTB_A125_2.EOF 
	
		sHTFBody =  sHTFBody & "85"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
        '3. 페이지번호 
        If Not ChkNotNull(lgcTB_A125_2.GetData("SEQ_NO"), "페이지번호") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("SEQ_NO"), 4, 0)

        '4.일련번호 
         If Not ChkNotNull(lgcTB_A125_2.GetData("SEQ"), "일련번호") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("SEQ"), 6, 0)
		
        '5.자회사명 
         If Not ChkNotNull(lgcTB_A125_2.GetData("W1"), "자회사명") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W1"), 50)
		
        '6.업종코드 
        
			Call SubMakeSQLStatements_W9125MA1("2",lgcTB_A125_2.GetData("W2"), "", "")  '업종체크 
			If   FncOpenRs("R",lgObjConn, oRs3, lgStrSQL, "", "") = False Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_A125_2.GetData("W2"), UNIGetMesg("업종코드가 존재하지 않습니다.", "",""))
			End If
		
		
         If Not ChkNotNull(lgcTB_A125_2.GetData("W2"), "업종코드") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W2"), 7)
		
        '7 소재지 
          If Not ChkNotNull(lgcTB_A125_2.GetData("W3"), "소재지") Then blnError = True	
		 sHTFBody = sHTFBody & UNIChar(lgcTB_A125_2.GetData("W3"), 70)


        '8 현지법인의출자금액 
         If Not ChkNotNull(lgcTB_A125_2.GetData("W4"), "현지법인의출자금액") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W4"), 15, 0)
		 
        '9 출자비율 
        If Not ChkNotNull(lgcTB_A125_2.GetData("W5"), "출자비율") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W5"), 5, 1)
		 
        '10 당기순이익 
        If Not ChkNotNull(lgcTB_A125_2.GetData("W6"), "당기순이익") Then blnError = True	
		 sHTFBody = sHTFBody & UNINumeric(lgcTB_A125_2.GetData("W6"), 15, 0)
		 
        '11 공란 22
         sHTFBody = sHTFBody & UNIChar("", 22)  & vbCrLf' -- 공란	 :
     
		lgcTB_A125_2.MoveNext 
	
	Loop
	
	'sHTFBody = mid(sHTFBody, 1,len(sHTFBody)- 1)

	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
		
	End If
	Call SubCloseRs(oRs3)

	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_A125 = Nothing	' -- 메모리해제 
	Set lgcTB_A125_1 = Nothing	' -- 메모리해제 
	Set lgcTB_A125_2 = Nothing	' -- 메모리해제 
	
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9125MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
		Case "A126" '-- 외부 참조 SQL

			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT TOP 1 1 FROM TB_A126 A "  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W01 > 0 "	 & vbCrLf	' -- 자산총계 
			

		Case "1"
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT TOP 1 1 FROM TB_A125 A"  & vbCrLf
'			lgStrSQL = lgStrSQL & "		INNER JOIN TB_A125 B ON A."  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.W01 > 0 "	 & vbCrLf	' -- 자산총계 
		
		
		Case "2" '업종체크 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " select top 1 STD_INCM_RT_CD from tb_std_income_rate"  & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE ATTRIBUTE_YEAR = '2005' " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND STD_INCM_RT_CD= " & filterVar(pCode1 ,"''","S")	 & vbCrLf
	
	
	End Select
	PrintLog "SubMakeSQLStatements_W9125MA1 : " & lgStrSQL
End Sub
%>
