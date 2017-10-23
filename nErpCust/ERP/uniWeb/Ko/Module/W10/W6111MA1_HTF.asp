<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 조특제3호연구및인력개발명세서 
'*  3. Program ID           : W6111MA1
'*  4. Program Name         : W6111MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_JT3

Set lgcTB_JT3 = Nothing ' -- 초기화 

Class C_TB_JT3
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.

	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' --서식을 읽어온다.
		Call SubMakeSQLStatements("A",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("B",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		If blnData1 = False And blnData2 = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
	On Error Resume Next   

		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
					
				End If
		End Select
		
		if err.number <> 0 then
					   Response.Write pFieldNm
					   Response.End 
					End if 
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
		Call SubCloseRs(lgoRs2)		
	End Sub

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_JT3A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "B"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_JT3B	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6111MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6111MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6111MA1"

	Set lgcTB_JT3 = New C_TB_JT3		' -- 해당서식 클래스 
	
	If Not lgcTB_JT3.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W6111MA1

	'==========================================
	' -- 조특제3호연구및인력개발명세서 오류검증 
	' -- 1. 매출및매입거래등 
	'==========================================
	iSeqNo = 1	
	'Response.End 'zzz
	Do Until lgcTB_JT3.EOF(2) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_JT3.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "ACCT_NM"), "계정과목") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W1_T"), "구분및과목(6)") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W2_T"), "구분및과목(7)") Then blnError = True
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W3_T"), "구분및과목(8)") Then blnError = True	
			 If Not ChkNotNull(lgcTB_JT3.GetData(2, "W4_T"), "구분및과목(9)") Then blnError = True	
			 
			 
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "SEQ_NO"), 6)
			
			  
		End If
					
		 If UNICDbl(lgcTB_JT3.GetData(2, "W6"),0) <> UNICDbl(lgcTB_JT3.GetData(2, "W1"),0) + UNICDbl(lgcTB_JT3.GetData(2, "W2"),0) + UNICDbl(lgcTB_JT3.GetData(2, "W3"),0)+ UNICDbl(lgcTB_JT3.GetData(2, "W4"),0)+ UNICDbl(lgcTB_JT3.GetData(2, "W5"),0) Then
		   '구분및비목의 항목(6)금액 + (7) + (8) + (9) + (10)의 금액의 계 
		    Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_JT3.GetData(2, "W6"),0) ,UNIGetMesg(TYPE_CHK_NOT_EQUAL, "계정" & lgcTB_JT3.GetData(2, "ACCT_NM")  & "합계","각 구분 비목의 합"))
			blnError = True	
		 End If
	

		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "ACCT_NM"), 20)
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W1_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W1"), "금액(6)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W1"), 15, 0)

			
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W2_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W2"), "금액(7)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W2"), 15, 0)

		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W3_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W3"), "금액(8)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W3"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(2, "W4_T"), 20)
		
		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W4"), "금액(9)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W4"), 15, 0)

		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W5"), "금액(10)") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W5"), 15, 0)

		If Not ChkNotNull(lgcTB_JT3.GetData(2, "W6"), "금액_") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(2, "W6"), 15, 0)

		sHTFBody = sHTFBody & UNIChar("", 48) & vbCrLf	' -- 공란 

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_JT3.MoveNext(2)	' -- 2번 레코드셋 
	Loop

	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""

	'==========================================
	' 2 연구및인력개발비명세서_연구및인력개발비의증가발생액의계산,공제세액 
	'==========================================
	
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D1_S"), "직전4년간발생합계액_기간1시작년월일") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D1_S"), "직전4년간발생합계액_기간1시작년월일") Then
	       blnError = True	
	    End if
	   
	Else
	    blnError = True	
	End If    

	  	 sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D1_S"))
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D1_E"), "직전4년간발생합계액_기간1종료년월일") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D1_E"), "직전4년간발생합계액_기간1종료년월일") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D1_E"))
			  
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT1"), "직전4년간발생합계액_기간1금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT1"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D2_S"), "직전4년간발생합계액_기간2시작년월일") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D2_S"), "직전4년간발생합계액_기간2시작년월일") Then
	       blnError = True	
	    End if
	Else
		blnError = True	
	End if	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D2_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D2_E"), "직전4년간발생합계액_기간2종료년월일") Then 
	   If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D2_E"), "직전4년간발생합계액_기간2종료년월일") Then
	       blnError = True	
	    End if
	   
	Else
	    blnError = True	
	End if   
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D2_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT2"), "직전4년간발생합계액_기간2금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT2"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D3_S"), "직전4년간발생합계액_기간3시작년월일") Then
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D3_S"), "직전4년간발생합계액_기간3시작년월일") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D3_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D3_E"), "직전4년간발생합계액_기간3종료년월일") Then
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D3_E"), "직전4년간발생합계액_기간3종료년월일") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D3_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT3"), "직전4년간발생합계액_기간3금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT3"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D4_S"), "직전4년간발생합계액_기간4시작년월일") Then
	     If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D4_S"), "직전4년간발생합계액_기간4시작년월일") Then
	       blnError = True	
	    End if
	    
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D4_S"))
		
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_D4_E"), "직전4년간발생합계액_기간4종료년월일") Then 
	    If Not ChkDate(lgcTB_JT3.GetData(1, "W8_D4_E"), "직전4년간발생합계액_기간4종료년월일") Then
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_JT3.GetData(1, "W8_D4_E"))
			
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W8_AMT4"), "직전4년간발생합계액_기간4금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_AMT4"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W8_SUM"), "직전4년간발생합계액_계") Then 
	   If UNICDbl(lgcTB_JT3.GetData(1, "W8_SUM"),0) <> UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT1"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT2"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT3"),0) + UNICDbl(lgcTB_JT3.GetData(1, "W8_AMT4"),0) Then
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W8_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "직전4년간발생합계액_계","각 기간별의 금액의 합"))
			blnError = True	
	   End If
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W8_SUM"), 15, 0)

	'항목(14)직전4년간 연편균발생액 = 항목(13)직전4년간발생합계액_계 X (48/직전4년간의사업연도월수) X (1/4)X (당해연도월수/12)
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W9"), "직전4년간연평균발생") Then
	    If Unicdbl(lgcTB_JT3.GetData(1, "W9"),0) <> fix(Unicdbl(lgcTB_JT3.GetData(1, "W8_SUM"),0) * (48/Unicdbl(lgcTB_JT3.GetData(1, "W_4Year_Mth"),0))*0.25* (Unicdbl(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"),0)/12) ) Then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W9"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "직전4년간 연편균발생액","직전4년간발생합계액_계 X (48/직전4년간의사업연도월수) X (1/4)X (당해연도월수/12)"))
	    	blnError = True	
	    End if
	Else
	   blnError = True	
	End if   
	 

	sHTFBody = sHTFBody &  UNINumeric(lgcTB_JT3.GetData(1, "W9"), 15, 0)

	If  ChkNotNull(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), "직전4년간의사업연도월수") Then 
	    If UNICDbl(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"),0) > 48 then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "직전4년간의사업연도월수","48"))
	       blnError = True	
	    End if
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W_4YEAR_Mth"), 2, 0)
	
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), "당해연도월수") Then 
	   If UNICDbl(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"),0) > 12 then
	  
	   
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "당해연도월수","12"))
	       blnError = True	
	    End if
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W_CURR_YEAR_Mth"), 2, 0)

    '항목(15)증가발생금액	= 항목(12)당해연도의연구및인력개발비발생명세의 금액(계) 합계 -항목(14)직전4년간연평균발생액 
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W10"), "증가발생금액") Then
	   ' IF UNICDbl(lgcTB_JT3.GetData(1, "W10"),0) <>  UNICDbl(lgcTB_JT3.GetData(1, "W15_11"),0) - Unicdbl(lgcTB_JT3.GetData(1, "W9"),0) Then
	    '   Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "증가발생금액","당해연도의연구및인력개발비발생명세의 금액(계) 합계 - 직전4년간연평균발생액"))
	     '  blnError = True	
	   ' End if

	Else
	     blnError = True	
	End If     
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W10"), 15, 0)
	'	
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_11"), "당해연도총발생금액공제_대상금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W15_11"), 15, 0)

	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_12_VALUE"), "당해연도총발생금액공제_공제율") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(UNICDbl(lgcTB_JT3.GetData(1, "W15_12_VALUE"),0) * 100, 5, 2)
		
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W15_13"), "당해연도총발생금액공제_공제세액") Then blnError = True	
	
	'항목(20)의(18)당해연도총발생금액공제_공제세액	=항목(20)의 (16)당해연도총발생금액공제_대상금액 X 항목(20)의 (17)당해연도총발생금액공제_공제율(15/100)(허용오차±10000)	
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W15_13"), "당해연도총발생금액공제_공제세액") Then
	    sTmp =  UNICDbl(lgcTB_JT3.GetData(1, "W15_11"),0) * UNICDbl(lgcTB_JT3.GetData(1, "W15_12_Value"),0) 
	    If (sTmp - 10000) <=  Unicdbl(lgcTB_JT3.GetData(1, "W15_13"),0)  And  Unicdbl(lgcTB_JT3.GetData(1, "W15_13"),0)  <= (sTmp + 10000)  Then
	    Else
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W15_13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당해연도총발생금액공제_공제세액","당해연도총발생금액공제_대상금액 * 당해연도총발생금액공제_공제율"))
	    End if
	Else
	    blnError = True	
	End If    
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W15_13"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W15_14"), 30)'"당해연도총발생금액공제_비고 
    
    '항목(21)의(16)증가발생금액공제_대상금액	= 항목(15)증가발생금액 
	If  ChkNotNull(lgcTB_JT3.GetData(1, "W16_11"), "증가발생금액공제_대상금액") Then 
	
	    If UNICDbl( lgcTB_JT3.GetData(1, "W16_11"),0) <> UNICDbl( lgcTB_JT3.GetData(1, "W10"),0) Then
	       Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W16_11"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "증가발생금액공제_대상금액","증가발생금액"))
	       blnError = True
	    End if   
	   
	Else
	   blnError = True	
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W16_11"), 15, 0)
		
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W16_12_Value"), "증가발생금액공제_공제율") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(UNICDbl(lgcTB_JT3.GetData(1, "W16_12_Value"),0) * 100, 5, 2)
	
	'항목(21)의(18)당해연도총발생금액공제_공제세액	- 항목(21)의 (16)당해연도총발생금액공제_대상금액 X 항목(21)의   (17)증가발생금액공제_공제율(40/100,중소기업의경우 50/100)(허용오차±10000)


	If  ChkNotNull(lgcTB_JT3.GetData(1, "W16_13"), "증가발생금액공제_공제세액") Then
	    sTmp =  UNICDbl(lgcTB_JT3.GetData(1, "W16_11"),0) * UNICDbl(lgcTB_JT3.GetData(1, "W16_12_Value"),0) 
	    If (sTmp - 10000) <=  Unicdbl(lgcTB_JT3.GetData(1, "W16_13"),0)  And  Unicdbl(lgcTB_JT3.GetData(1, "W16_13"),0)  <= (sTmp + 10000)  Then
	    Else
	        Call SaveHTFError(lgsPGM_ID,lgcTB_JT3.GetData(1, "W16_13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당해연도총발생금액공제_공제세액","당해연도총발생금액공제_대상금액 * 당해연도총발생금액공제_공제율"))
	    End if
	Else
	    blnError = True	
	End If    
	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W16_13"), 15, 0)


	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W16_14"), 30)   '증가발생금액공제_비고 
  
    '항목(22)의(18)당해연도에공제받을세액_공제세액	- 중소기업(항목(20)의 (18)당해연도총발생금액공제_공제세액과 항목(21)의   (18)증가발생금액공제_공제세액중 선택) 
    '또는 중소기업외의중소기업(항목(21)  의 (18) 증가발생금액공제_공제세액)  세액공제신청서(A165)의 코드(32) 연구 및 인력개발비 세액공제 항목(11) 대  상세액과 일치 
    ' (항목(22)의(18)당해연도에공제받을세액_공제세액이 “0”보다 큰 경우 반드시 입력)
	
	If Not ChkNotNull(lgcTB_JT3.GetData(1, "W17"), "당해연도에공제받을세액_공제세액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT3.GetData(1, "W17"), 15, 0)

	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT3.GetData(1, "W17_14"), 60) '당해연도에공제받을세액_비고 


	sHTFBody = sHTFBody & UNIChar("", 16) & vbCrLf	' -- 공란 

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_JT3 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL
			

	End Select
	PrintLog "SubMakeSQLStatements_W6111MA1 : " & lgStrSQL
End Sub
%>
