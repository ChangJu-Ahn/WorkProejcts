<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제55호 소득자료 명세서 
'*  3. Program ID           : W9113MA1
'*  4. Program Name         : W9113MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_55

Set lgcTB_55 = Nothing ' -- 초기화 

Class C_TB_55
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
			GetData = lgoRs1(pFieldNm)
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
				lgStrSQL = lgStrSQL & " FROM TB_55	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9113MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9113MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum, dblw4_Sum, dblw5_Sum ,dblw4_99, dblw5_99
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9113MA1"

	Set lgcTB_55 = New C_TB_55		' -- 해당서식 클래스 
	
	If Not lgcTB_55.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9113MA1

	'==========================================
	' --제55호 소득자료 명세서 오류검증 
	iSeqNo = 1	: sHTFBody = ""
	dblw4_Sum = 0
	dblw5_Sum = 0
	Do Until lgcTB_55.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_55.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			'인정상여는 ‘1’, 인정배당은 ‘2’, 기타소득은 ‘3’	
			If Not ChkBoundary("1,2,3", lgcTB_55.GetData("W1") , "인정상여코드" ) Then  blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W1"), "소득구분") Then blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W2"), "사업연도") Then blnError = True
			'If Not ChkDate(lgcTB_55.GetData("W2"),"소득구분코드" & lgcTB_55.GetData("W1") & "사업연도") Then  blnError = True	
			If Not ChkNotNull(lgcTB_55.GetData("W3"), "소득귀속연도") Then blnError = True
			'If Not ChkDate(lgcTB_55.GetData("W3"),"소득구분코드" & lgcTB_55.GetData("W1") & "귀속연도") Then  blnError = True	
		
			
		    If  Len(UNIRemoveDash(lgcTB_55.GetData("W8"))) <> 10  AND Len(UNIRemoveDash(lgcTB_55.GetData("W8")) ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_55.GetData("W8"), UNIGetMesg(TYPE_CHK_CHARNUM, "소득구분코드" & lgcTB_55.GetData("W1") & "주민등록번호","10 이거나 13"))
				blnError = True	
			End If
			  dblw5_Sum = dblw5_Sum  + unicdbl(lgcTB_55.GetData("W5"),0)
			  dblw4_Sum = dblw4_Sum  + unicdbl(lgcTB_55.GetData("W4"),0)
			 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("SEQ_NO"), 6)
			dblw4_99 = Unicdbl(lgcTB_55.GetData("W4"),0)
			dblw5_99 = Unicdbl(lgcTB_55.GetData("W5"),0)
		End If
		
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W1"), 1)
				
		sHTFBody = sHTFBody & UNI8Date(lgcTB_55.GetData("W2"))
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W3"), 4)
	
		If  ChkNotNull(lgcTB_55.GetData("W4"), "배당.상여 및 기타소득금액") Then 
		   
		Else
			blnError = True
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_55.GetData("W4"), 15, 0)
		
		If  Not ChkNotNull(lgcTB_55.GetData("W5"), "원천징수할소득금액") Then   blnError = True
		
	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_55.GetData("W5"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W6"), 70)
		sHTFBody = sHTFBody & UNIChar(lgcTB_55.GetData("W7"), 30)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_55.GetData("W8")), 13)
		sHTFBody = sHTFBody & UNIChar("", 32) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		lgcTB_55.MoveNext 
	Loop
	
	
	  If dblw4_Sum <> dblw4_99 Then 
						
	   Call SaveHTFError(lgsPGM_ID,dblw4_Sum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "배당.상여 및 기타소득금액합계","각 배당.상여 및 기타소득금액 합"))
	   blnError = True	
	 End if
	
	 If dblw5_Sum <> dblw5_99 Then 
						
	   Call SaveHTFError(lgsPGM_ID,dblw5_Sum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "원천징수할소득금액합계","각 원천징수할소득금액 합"))
	   blnError = True	
	End if
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_55 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9113MA1 : " & lgStrSQL
End Sub
%>
