<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제50호 자본금과 적립금조정명세서(을)
'*  3. Program ID           : W7109MA1
'*  4. Program Name         : W7109MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_50B

Set lgcTB_50B = Nothing ' -- 초기화 

Class C_TB_50B
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
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
	            If WHERE_SQL = "" Then 
					lgStrSQL = lgStrSQL & " , ( SELECT ITEM_NM FROM TB_ADJUST_ITEM WITH (NOLOCK)  WHERE ITEM_CD = A.W1 ) W1_NM " & vbCrLf
				Else
					lgStrSQL = lgStrSQL & " , '' W1_NM  " & vbCrLf
				End If
				lgStrSQL = lgStrSQL & " FROM TB_50B	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W7109MA1
	Dim A102

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W7109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo,sTmp1,sTmp2
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W7109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W7109MA1"

	Set lgcTB_50B = New C_TB_50B		' -- 해당서식 클래스 
	
	If Not lgcTB_50B.LoadData Then Exit Function			' -- 제50호 자본금과 적립금조정명세서(을) 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W7109MA1

	'==========================================
	' -- 제15호 소득금액조정합계표 오류검증 
	iSeqNo = 1	: sHTFBody = ""
	
	Do Until lgcTB_50B.EOF() 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_50B.GetData("SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			
			If Not ChkNotNull(lgcTB_50B.GetData("W1_NM"), "과목 또는 사항") Then blnError = True	
			
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("SEQ_NO"), 6)
			
			' -- 합계일때 검증실시 
			' -- 서식 15호를 로드한다.
			Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp 에 정의됨 
				
			' -- 추가 조회조건을 읽어온다.  400합 
			Call SubMakeSQLStatements_W7109MA1("A102_1",iKey1, iKey2, iKey3)   
				
			cDataExists.A102.CALLED_OUT	= True				' -- 외부에서 호출함을 알림 
			cDataExists.A102.WHERE_SQL	= lgStrSQL			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			cDataExists.A102.SELECT_SQL	= " W3, SUM(W2) W2 "' -- 다른 리턴 내용 
			
			
			If Not cDataExists.A102.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제15호 과목별소득금액조정명세서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				
		   '소득금액조정합계표(A102)의 익금산입 처분코드 ‘400’의 금액합 - 손금산입 처분코드 ‘100’의 금액합 = 항목(4)당기중 증가의 합 - 항목(3)당기중 감소의 합		
				'- 소득금액조정합계표(A102)의 손금산입 처분코드 ‘100:유보’의 금액합 
				sTmp1 =  UNICDbl(cDataExists.A102.GetData("W2"), 0) 
				
				Call cDataExists.A102.MoveNext 
				
				'- 소득금액조정합계표(A102)의 익금산입 처분코드 ‘400:유보’의 금액합 
				sTmp2 =  UNICDbl(cDataExists.A102.GetData("W2"), 0) 
				If UNICDbl(lgcTB_50B.GetData("W4"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0)  <> UNICDbl(sTmp2, 0) -  UNICDbl(sTmp1, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_50B.GetData("W4"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0)  , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당기중 증가의 합 - 당기중 감소의 합","소득금액조정합계표(A102)의 익금산입 처분코드 ‘400’의 금액합 - 손금산입 처분코드 ‘100’의 금액합"))
				End If
			End If
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A102 = Nothing

		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("W1_NM"), 40)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W2"), lgcTB_50B.GetData("W1_NM") & "_기초잔액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W2"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W3"), lgcTB_50B.GetData("W1_NM") & "_당기중감소") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W3"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W4"), lgcTB_50B.GetData("W1_NM") & "_당기중증가") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_50B.GetData("W5"), lgcTB_50B.GetData("W1_NM") & "_기말잔액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_50B.GetData("W5"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_50B.GetData("W_DESC"), 50) ' -- 널허용 
	
		' -- 기말잔액 = 기초잔액 - 당기중감소 + 당기중증감 
		If UNICDbl(lgcTB_50B.GetData("W2"), 0) - UNICDbl(lgcTB_50B.GetData("W3"), 0) + UNICDbl(lgcTB_50B.GetData("W4"), 0) <> UNICDbl(lgcTB_50B.GetData("W5"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기말잔액","기초잔액 - 당기중감소 + 당기중증감"))
		End If
	
		sHTFBody = sHTFBody & UNIChar("", 38) & vbCrLf	' -- 공란 
		
		iSeqNo = iSeqNo + 1
		
		lgcTB_50B.MoveNext 
	Loop

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_50B = Nothing	' -- 메모리해제 

End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W7109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A102_1" '-- 외부 참조 SQL
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND (  (A.W_TYPE	= '1' AND A.W3 = '400')  " 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		OR (A.W_TYPE	= '2' AND A.W3 = '100') )" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	GROUP BY A.W3 " 	 & vbCrLf

	End Select
	PrintLog "SubMakeSQLStatements_W7109MA1 : " & lgStrSQL
End Sub
%>
