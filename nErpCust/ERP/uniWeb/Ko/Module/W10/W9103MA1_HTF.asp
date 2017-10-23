<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제47호 주요계정명세서(을)
'*  3. Program ID           : W9103MA1
'*  4. Program Name         : W9103MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_47B

Set lgcTB_47B = Nothing ' -- 초기화 

Class C_TB_47B
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		
	Private lgoRs3		
	Private lgoRs4		
	
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3, blnData4
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' --서식을 읽어온다.
		Call SubMakeSQLStatements("1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
			Exit Function
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
			
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("3",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("4",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs4,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData4 = False
		End If
		
		If blnData1 = False And blnData2 = False And blnData3 = False And blnData4 = False Then
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
			Case 3
				lgoRs3.Find pWhereSQL
			Case 4
				lgoRs4.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
			Case 2
				lgoRs2.Filter = pWhereSQL
			Case 3
				lgoRs3.Filter = pWhereSQL
			Case 4
				lgoRs4.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
			Case 3
				EOF = lgoRs3.EOF
			Case 4
				EOF = lgoRs4.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
			Case 3
				lgoRs3.MoveFirst
			Case 4
				lgoRs4.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
			Case 3
				lgoRs3.MoveNext
			Case 4
				lgoRs4.MoveNext
		End Select
	End Function	
	
	Function GetData(Byval pType, Byval pFieldNm)
		Select Case pType
			Case 1
				If Not lgoRs1.EOF Then
					GetData = lgoRs1(pFieldNm)
				End If
			Case 2
				If Not lgoRs2.EOF Then
					GetData = lgoRs2(pFieldNm)
				End If
			Case 3
				If Not lgoRs3.EOF Then
					GetData = lgoRs3(pFieldNm)
				End If
			Case 4
				If Not lgoRs4.EOF Then
					GetData = lgoRs4(pFieldNm)
				End If
		End Select
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
		Call SubCloseRs(lgoRs4)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
		Call SubCloseRs(lgoRs4)	
	End Sub

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT  "
			lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
			lgStrSQL = lgStrSQL & " FROM TB_47B" & pMode & "	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
					
			If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		
	 

			PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9103MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9103MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists ,sTmp1,Stmp2
    Dim iSeqNo, sMsg
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9103MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9103MA1"

	Set lgcTB_47B = New C_TB_47B		' -- 해당서식 클래스 
	
	If Not lgcTB_47B.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9103MA1

	'==========================================
	' -- 제47호 주요계정명세서(을) 오류검증 
	' -- 1. 매출및매입거래등 
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
	Do Until lgcTB_47B.EOF(1) 
	
		sTmp = lgcTB_47B.GetData(1, "W1")
		
		Select Case sTmp
			Case "101"
				sMsg = "제품및상품"
			Case "102"
				sMsg = "반제품및재공품"
			Case "103"
				sMsg = "원재료"
			Case "104"
				sMsg = "저장품"
			Case "105"
				sMsg = "유가증권_채권"
			Case "106"
				sMsg = "유가증권_기타"
			Case "107"
				sMsg = "합계"
		End Select
		
		
		
		
		
		 If sMsg = "101" Or  sMsg = "102"   Or  sMsg = "103" Or  sMsg = "104" Or  sMsg = "105" Then
		     sTmp1 ="1,2,3,4,5,6,7,8"
		  	  	'1:개별법, 2:선입선출법, 3:후입선출법, 4:총평균법, 5:이동평균법, 6:매출가격환원법 7:저가법, 8:기타 
		 Else	
		    '1:개별법, 2:총평균법, 3:이동평균법, 4:시가법,5:기타 
		    sTmp1 ="1,2,3,4,5,"
		 
		 End If
			   
			   
	     IF sTmp <> "107"  AND  UNICDbl(lgcTB_47B.GetData(1, "W4"),0)  <> 0 Then 
		   
		     If  ChkNotNull(lgcTB_47B.GetData(1, "W2"), sMsg & "_신고방법") Then
		       
		         IF Not ChkBoundary(sTmp1, lgcTB_47B.GetData(1, "W2"),sMsg & "_신고방법") Then
		 	       blnError = True	
		 	       
		 	    End if
		 	Else
		 	    blnError = True	
		 	End IF   
		       
		    
		   
		   If  ChkNotNull(lgcTB_47B.GetData(1, "W3"), sMsg & "_평가방법") Then 
				IF Not ChkBoundary(sTmp1, lgcTB_47B.GetData(1, "W3"),sMsg & "_평가방법") Then
				       blnError = True	
				
				End IF   
		   End if  
		   
		  
	  
	
		End IF   
	
		IF sTmp <> "107" Then
		    sHTFBody = sHTFBody & UNIChar(lgcTB_47B.GetData(1, "W2"), 1) '신고방법 
				
		   sHTFBody = sHTFBody & UNIChar(lgcTB_47B.GetData(1, "W3"), 1) '평가방법 
		End If 
		

		If Not ChkNotNull(lgcTB_47B.GetData(1, "W4"), sMsg & "_회사계산금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W4"), 15, 0)
		
	
		If Not ChkNotNull(lgcTB_47B.GetData(1, "W5"), sMsg & "_조정계산금액_신고방법") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W5"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(1, "W6"),  sMsg & "_조정계산금액_선입선출법") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W6"), 15, 0)
		
		'-조정액	- 항목(5)조정계산금액_신고방법 - 항목(4)회사계산금액또는 (항목(5)와 항목(6)조정계산금액_선입선출법 중 큰 금액-항목(4)회사계산금액)
		If  ChkNotNull(lgcTB_47B.GetData(1, "W7"), sMsg & "_조정액") Then 
		
		     If UniCDBL(lgcTB_47B.GetData(1, "W5"),0) > UNICDbl(lgcTB_47B.GetData(1, "W6"),0)  Then
		        Stmp2 = UNICDBl(lgcTB_47B.GetData(1, "W5"),0)  - UNICDbl(lgcTB_47B.GetData(1, "W4"),0)
		     Else
		        Stmp2 =UNICDbl(lgcTB_47B.GetData(1, "W6"),0) - UNICDbl(lgcTB_47B.GetData(1, "W4"),0)
		     End if
		    IF  UniCDBL(lgcTB_47B.GetData(1, "W7"),0)  <> UNICDBl(lgcTB_47B.GetData(1, "W5"),0) - UNICDbl(lgcTB_47B.GetData(1, "W5"),4)  Or   UniCDBL(lgcTB_47B.GetData(1, "W7"),0) <> Stmp2 Then
		        blnError = True	
		        
			
		       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(1, "W7"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "조정액"," 조정계산금액_신고방법 - 항목(4)회사계산금액 또는 (항목(5)와 항목(6)조정계산금액_선입선출법 중 큰 금액-항목(4)회사계산금액  "))
		    End If
		
		Else
		   blnError = True	
		End If  
		
		
 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(1, "W7"), 15, 0)
		

		Call lgcTB_47B.MoveNext(1)	' -- 1번 레코드셋 
	Loop

	' 2. 그리드2
	Do Until lgcTB_47B.EOF(2) 
	
		sTmp = lgcTB_47B.GetData(2, "W8")
		
		Select Case sTmp
			Case "108"
				sMsg = "국고보조금"
			Case "109"
				sMsg = "공사부담금"
			Case "110"
				sMsg = "보험차익"
		End Select
		

		
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W9"), sMsg & "_금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W9"), 15, 0)
				
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W10"), sMsg & "_취득고정자산가액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W11"),  sMsg & "_회사손금계상액") Then blnError = True	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W11"), 15, 0)
		
		If  ChkNotNull(lgcTB_47B.GetData(2, "W12"), sMsg & "_한도초과액") Then 
		    '- 항목(11)회사손금계상액 - 항목(10)취득고정자산가액 
		   IF UniCDBL(lgcTB_47B.GetData(2, "W12"),0)  <>  UNICDBl(lgcTB_47B.GetData(2, "W11"),0)   -UNICDBl(lgcTB_47B.GetData(2, "W10"),0)  Then
		      Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(2, "W12"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_한도초과액", "(11)회사손금계상액 - 항목(10)취득고정자산가액"))
		      blnError = True	
		
		   End If
		Else
		    blnError = True	
		End If    
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W12"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(2, "W13"), sMsg & "_미사용분익금산입액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(2, "W13"), 15, 0)
		
		Call lgcTB_47B.MoveNext(2)	' -- 2번 레코드셋 
	Loop
	
	' 3. 그리드 3
	
		
			

	If Not ChkNotNull(lgcTB_47B.GetData(3, "W14"), "가지급금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W14"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W15"), "가수금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W15"), 15, 0)
				
	If  ChkNotNull(lgcTB_47B.GetData(3, "W16"), "차감") Then 
	   '적수 차감	- 항목(14)가지급금 - 항목(15)가수금 
	    IF UNICDBl(lgcTB_47B.GetData(3, "W16"),0) <> UNICDBl(lgcTB_47B.GetData(3, "W14"),0) -UNICDBl(lgcTB_47B.GetData(3, "W15"),0) Then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(3, "W16"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_차감", "항목(14)가지급금 - 항목(15)가수금"))
	
	    End If
	Else
	   blnError = True	
	End If
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W16"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W17"), "인정이자") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W17"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W18"), "회사계상액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W18"), 15, 0)
				
	If Not ChkNotNull(lgcTB_47B.GetData(3, "W19"), "조정액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(3, "W19"), 15, 0)
				
	
	' 4. 그리드4
	Do Until lgcTB_47B.EOF(4) 
	
		sTmp = lgcTB_47B.GetData(4, "W20")
		
		Select Case sTmp
			Case "111"
				sMsg = "건설완료자산분"
			Case "112"
				sMsg = "건설중인자산분"
			Case "113"
				sMsg = "계"
		End Select
		
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W21"), sMsg & "_건설자금이자") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W21"), 15, 0)
				
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W22"), sMsg & "_회사계상액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W22"), 15, 0)
		
		If Not ChkNotNull(lgcTB_47B.GetData(4, "W23"),  sMsg & "_상각대상자산분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W23"), 15, 0)
		
		''차감조정액_건설완료자산분	= 항목(21)건설자금이자 - 항목(22)회사계상액 - 항목(23)상각대상자분		
		If  ChkNotNull(lgcTB_47B.GetData(4, "W24"), sMsg & "_차감조정액") Then 
		    IF UNICDbl(lgcTB_47B.GetData(4, "W24"),0) <> UNIcdbl(lgcTB_47B.GetData(4, "W21"),0)-UNIcdbl(lgcTB_47B.GetData(4, "W22"),0) - UNICDbl(lgcTB_47B.GetData(4, "W23"),0) Then
		       Call SaveHTFError(lgsPGM_ID, lgcTB_47B.GetData(3, "W24"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, sMsg & "_차감", "항목(21)건설자금이자 - 항목(22)회사계상액 - 항목(23)상각대상자분"))
		       blnError = True	 
		    End IF
		    
		    
		Else
			blnError = True	
		End IF	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_47B.GetData(4, "W24"), 15, 0)

		Call lgcTB_47B.MoveNext(4)	' -- 2번 레코드셋 
	Loop
	
	sHTFBody = sHTFBody & UNIChar("", 117) & vbCrLf	' -- 공란 

	PrintLog "WriteLine2File : " & sHTFBody

 
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_47B = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9103MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9103MA1 : " & lgStrSQL
End Sub
%>
