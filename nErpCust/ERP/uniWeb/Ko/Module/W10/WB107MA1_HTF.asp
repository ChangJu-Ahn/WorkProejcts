<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제51호 중소기업기준검토표 
'*  3. Program ID           : WB107MA1
'*  4. Program Name         : WB107MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_51

Set lgcTB_51 = Nothing	' -- 초기화 

Class C_TB_51
	' -- 테이블의 컬럼변수 
	
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
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
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
				lgStrSQL = lgStrSQL & " FROM TB_51	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_WB107MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_WB107MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_WB107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "WB107MA1"
	
	Set lgcTB_51 = New C_TB_51		' -- 해당서식 클래스 
	
	If Not lgcTB_51.LoadData	Then Exit Function		' -- 제1호 서식 로드 
		
	'==========================================
	' -- 제51호 중소기업기준검토표 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 


    If unicdbl(lgcTB_51.GetData("W07"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W01"), "사업수입금액1이 0이 아니면 업태1") Then blnError = True		
	End if	
		sHTFBody = sHTFBody & UNIChar(Trim(lgcTB_51.GetData("W01")), 30)

	If unicdbl(lgcTB_51.GetData("W07"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W04"), "사업수입금액1이 0이 아니면 기준경비율코드1") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W04"), 7)
	
	If Not ChkNotNull(lgcTB_51.GetData("W07"), "사업수입금액1") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W07"), 15, 0)
	
	If unicdbl(lgcTB_51.GetData("W08"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W02"), "사업수입금액2이 0이 아니면 업태2") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W02"), 30)
	

	
	If unicdbl(lgcTB_51.GetData("W08"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W05"), "사업수입금액2이 0이 아니면 기준경비율코드2") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W05"), 7)                 '기준경비율코드2
	
		
	If Not ChkNotNull(lgcTB_51.GetData("W08"), "사업수입금액2") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W08"), 15, 0)          ' 사업수입금액2
	
	If unicdbl(lgcTB_51.GetData("W09"),0) <> 0 then
		If Not ChkNotNull(lgcTB_51.GetData("W06"), "사업수입금액-기타가 0이 아니면 기준경비율코드_기타") Then blnError = True		
	End if	
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W06"), 7)                 '기준경비율코드-기타 
	
	If Not ChkNotNull(lgcTB_51.GetData("W09"), "사업수입금액-기타") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W09"), 15, 0)          ' 사업수입금액-기타 

	
	
	If ChkNotNull(lgcTB_51.GetData("W_SUM"), "계_사업수입금액") Then 
	   '계-사업수입금액 : 항목 (7) + (8) + (9)
	   If UNICDbl(lgcTB_51.GetData("W_SUM"),0)  <> unicdbl(lgcTB_51.GetData("W07"),0) + UniCdbl(lgcTB_51.GetData("W08"),0)+ UniCdbl(lgcTB_51.GetData("W09"),0) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "계_사업수입금액","사업수입금액합"))
	        blnError = True		
	   End if
	Else
		'blnError = True		
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W_SUM"), 15, 0)
	
	If Not ChkBoundary("1,2", lgcTB_51.GetData("W19"), "해당사업_적합여부" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W19"), 1)

	If Not ChkNotNull(lgcTB_51.GetData("W10"), "상시종업원수") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W10"), 4, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W11"), "중소기업기본법시행령 별표1의 규모기준_종업원수") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W11"), 4, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W12"), "자본금") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W12"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W13"), "중소기업기본법시행령 별표1의 규모기준_자본금") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W13"), 6, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W14"), "매출액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W14"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W15"), "중소기업기본법시행령 별표1의 규모기준_매출액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W15"), 6, 0)
	
	If Not ChkNotNull(lgcTB_51.GetData("W16"), "자기자본") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W16"), 7, 1)
	
	If Not ChkNotNull(lgcTB_51.GetData("W17"), "상장.협회등록법인의 경우 자산총액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W17"), 7, 1)
	
	If Not ChkBoundary("1,2", lgcTB_51.GetData("W20"), "규모_적합여부" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	If Not ChkBoundary("1,2", lgcTB_51.GetData("W21"), "경영_적합여부" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W21"), 1)


    If lgcTB_51.GetData("W18") <> "" Then 
		If  unicdbl(lgcTB_51.GetData("W18"),0 ) <> 0 and unicdbl(lgcTB_51.GetData("W18"),0 ) <= 2001 Then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_CHK_LOW_AMT, "초과년도","2001이거나 2001"))
		     blnError = True
		End if
	Else
		'blnError = True		
	End if		
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_51.GetData("W18"), 4, 0)  '초과연도 


	If Not ChkBoundary("1,2", lgcTB_51.GetData("W22"), "유예기간_적합여부" ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	If  ChkBoundary("1,2", lgcTB_51.GetData("W23"), "적정여부" ) Then
	'- 항목 (19), (20), (21), (22) 이 모두 1(적합) 인 경우에 한해서  '1' (적합) 
	'- 항목(20)이 '2'(부적합) 이고 항목(22)가 1(적합)인 경우도 '1' (적합)
	   if lgcTB_51.GetData("W23") <> "1" and lgcTB_51.GetData("W19") ="1" and  lgcTB_51.GetData("W20") = "1" and  lgcTB_51.GetData("W21") = "1" and  lgcTB_51.GetData("W22") = "1" then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_MSG_NORMAL_ERR, "적정여부",""))
	        blnError = True
	   
	   End if
	   
	   
	     if lgcTB_51.GetData("W23") <> "1" and lgcTB_51.GetData("W20") ="2" and  lgcTB_51.GetData("W22") = "1"  then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_51.GetData("W18"), UNIGetMesg(TYPE_MSG_NORMAL_ERR, "적정여부",""))
	        
	        blnError = True
	   
	     End if
	  
	   
	Else
	  '  blnError = True
	End if
	sHTFBody = sHTFBody & UNIChar(lgcTB_51.GetData("W20"), 1)

	sHTFBody = sHTFBody & UNIChar("", 46)	' -- 공란 

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.

'zzzz  ??

blnError = false

	If Not blnError Then
		
		Call WriteLine2File(sHTFBody)
	End If
 
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_51 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_WB107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
	
	End Select
	PrintLog "SubMakeSQLStatements_WB107MA1 : " & lgStrSQL
End Sub

%>
