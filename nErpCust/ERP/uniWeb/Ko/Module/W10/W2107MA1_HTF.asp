
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제17호 조정후수입금액명세서 
'*  3. Program ID           : W2107MA1
'*  4. Program Name         : W2107MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_17

Set lgcTB_17 = Nothing ' -- 초기화 

Class C_TB_17
	' -- 테이블의 컬럼변수 
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs3		' -- 멀티로우 데이타는 지역변수로 선언한다.
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2, blnData3
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True	: blnData3 = False
		
		PrintLog "LoadData IS RUNNING: "
				 
		iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
		iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
		iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

		' --서식을 읽어온다.
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If

		' --서식을 읽어온다.
		Call SubMakeSQLStatements("D2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs3,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData3 = False
		End If
 
		If blnData1 = False And blnData2 = False And blnData3 = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If
		
		LoadData = True
	End Function

	'----------- 멀티 행 지원 ------------------------
'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
			    lgoRs1.Find pWhereSQL
		     Case 2
				lgoRs2.Find pWhereSQL
		     Case 3
				lgoRs3.Find pWhereSQL
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
		End Select
	End Function
	
	Function MoveFirst(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
			Case 3
				lgoRs3.MoveFirst
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
		End Select
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
		Call SubCloseRs(lgoRs3)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
		Call SubCloseRs(lgoRs2)		
		Call SubCloseRs(lgoRs3)
	End Sub


	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_17H A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_17_D1 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "     AND  ((A.W4 <> 0  and  A.W3 <>'') OR A.CODE_NO IN ( '11','99')) "  & vbCrLf

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
				lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_17_D2 A WITH (NOLOCK) " & vbCrLf	
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				lgStrSQL = lgStrSQL & "     AND A.W9 <> 0  "  & vbCrLf

				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W2107MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
	Dim A134
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W2107MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, iSeqNo
    Dim dblAmtW4, dblAmtW5, dblAmtW6, dblAmtW7
    Dim dblAmtSumW4, dblAmtSumW5, dblAmtSumW6, dblAmtSumW7
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W2107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W2107MA1"

	Set lgcTB_17 = New C_TB_17		' -- 해당서식 클래스 
	
	If Not lgcTB_17.LoadData Then Exit Function			
	
	Set cDataExists = new TYPE_DATA_EXIST_W2107MA1
	
	'==========================================
	' --제17호 조정후수입금액명세서 오류검증 
	iSeqNo = 1	
	
	Do Until lgcTB_17.EOF(1)
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

		If UNICDbl(lgcTB_17.GetData(1,"CODE_NO"), 0) <> "99" Then
			 
			 sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 
			 dblAmtW4 =  dblAmtW4 + UNICDbl(lgcTB_17.GetData(1,"W4"),0)
		     dblAmtW5 =  dblAmtW5 + UNICDbl(lgcTB_17.GetData(1,"W5"),0)
		     dblAmtW6 =  dblAmtW6 + UNICDbl(lgcTB_17.GetData(1,"W6"),0)
		     dblAmtW7 =  dblAmtW7 + UNICDbl(lgcTB_17.GetData(1,"W7"),0)
		     
		     Response.Write "dblAmtW4=" & dblAmtW4 & vbCrLf
		Else
			  sHTFBody = sHTFBody & UNIChar("999999", 6)
			 
			  dblAmtSumW4 =   UNICDbl(lgcTB_17.GetData(1,"W4"),0)
		      dblAmtSumW5 =   UNICDbl(lgcTB_17.GetData(1,"W5"),0)
		      dblAmtSumW6 =   UNICDbl(lgcTB_17.GetData(1,"W6"),0)
		      dblAmtSumW7 =   UNICDbl(lgcTB_17.GetData(1,"W7"),0)
		      '합비교 
		      
		      Response.Write "dblAmtSumW4=" & dblAmtSumW4 & vbCrLf
		End If
		
		
	if lgcTB_17.GetData(1,"W3") <> "" And UNICDbl(lgcTB_17.GetData(1,"W4"),0) <> 0 Then	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W1"), "업태")					Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W2"), "종목")					Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W3"), "기준(단순)경비율번호") Then blnError = True
		If Not ChkNotNull(lgcTB_17.GetData(1,"W4"), "수입금액_계")			Then blnError = True
			
		If Not ChkNotNull(lgcTB_17.GetData(1,"W5"), "수입금액_내수")		Then blnError = True	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W6"), "수입금액_수입상품")	Then blnError = True	
		If Not ChkNotNull(lgcTB_17.GetData(1,"W7"), "수입금액_수출")		Then blnError = True	
	End If	
			
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W1"), 30)			'업태 
				
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W2"), 30)			'종목 
				
		sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(1,"W3"), 7)			'기준(단순)경비율번호 
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W4"), 15, 0)	'수입금액_계 = 항목(5)국내생산품 + 항목 (6)수입상품 + 항목(7)수출 

	
		If  UNICDbl(lgcTB_17.GetData(1,"W4"),0) <> UNICDbl(lgcTB_17.GetData(1,"W5"),0) + UNICDbl(lgcTB_17.GetData(1,"W6"),0) +UNICDbl(lgcTB_17.GetData(1,"W7"),0)  Then
			 blnError = True
			 Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(1,"W4"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액_계","항목(5)국내생산품 + 항목 (6)수입상품 + 항목(7)수출 "))		
		End If
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W5"), 15, 0)	'수입금액_내수 
			
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W6"), 15, 0)	'수입금액_수입상품 
				
		sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(1,"W7"), 15, 0)	'수입금액_수출 
		 
		
		sHTFBody = sHTFBody & UNIChar("", 61) & vbCrLf	' -- 공란 
		
		
		lgcTB_17.MoveNext(1) 
		iSeqNo = iSeqNo + 1
	Loop

	If  dblAmtW4 <> dblAmtSumW4  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW4 & " <> " & dblAmtSumW4 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액_합계(코드 99)","각 수입금액_계의 합"))		
	End If
	
	If  dblAmtW5 <> dblAmtSumW5  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW5 & " <> " & dblAmtSumW5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액_내수(코드 99)","각 수입금액_내수의 합"))		
	End If
	
	If  dblAmtW6 <> dblAmtSumW6  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW6 & " <> " & dblAmtSumW6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액_수입상품(코드 99)","각 수입금액_수입상품의 합"))		
	End If
	
	If  dblAmtW7 <> dblAmtSumW7  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,dblAmtW7 & " <> " & dblAmtSumW7, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액_수출(코드 99)","수입금액_수출의 합"))		
	End If

	' -- 200603 개정:  조정후수입금액명세서 (A111)서식의 일련번호가 999999일때   - 수입금액조정명세서(A134)의 항목(6)조정후수입금액_계와 일치검증 추가 
	Set cDataExists.A134 = new C_TB_16	' -- W6127MA1_HTF.asp 에 정의됨 
											
	' -- 추가 조회조건을 읽어온다.
	cDataExists.A134.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
	cDataExists.A134.WHERE_SQL = " AND A.SEQ_NO = 999999 "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
											
	If Not cDataExists.A134.LoadData() Then
	
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "", lgcTB_5.W1 & "_대상세액이 '0'보다 큰 경우 최저한세적용계산서(A140) 서식 필수 입력 ")		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
	Else

		If UNICDbl(dblAmtSumW4, 0)  <> UNICDbl(cDataExists.A134.GetData(1, "W6") , 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, dblAmtSumW4 & " <> " & cDataExists.A134.GetData(1, "W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "제17호 조정후수입금액명세서(A111)서식의 코드(99)합계","제16호 수입금액조정명세서(A134)의 항목(6)조정후수입금액_계"))
		End If
													
	End If

	' -- 사용한 클래스 메모리 해제 
	Set cDataExists.A134 = Nothing		

	'------ 조정후 수입금액명세서_부가가치 과세표준과 수입금액차액검토 (헤더)
	iSeqNo = 1
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_17.GetData(2,"W8"), "부가가치세_과세표준_계") Then blnError = True	
    
    ' -- 200603 : 삭제됨 
	'sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W8"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W9"), "부가가치세_과세표준_일반") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W9"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W10"), "부가가치세_과세표준_영세율") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W10"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W11"), "면세사업수입금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W11"), 15, 0)
		
	If  ChkNotNull(lgcTB_17.GetData(2,"W12"), "합계") Then 
	    '항목(12)합계 : 항목(8)부가세과세표준_계 + 항목(11)면세사업수입금액 
	    IF UNICDbl(lgcTB_17.GetData(2,"W12"),0) <> UNICDbl(lgcTB_17.GetData(2,"W9"),0)  + UNICDbl(lgcTB_17.GetData(2,"W10"),0) + UNICDbl(lgcTB_17.GetData(2,"W11"),0)  Then
	       Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(2,"W12"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "합계","항목(8)부가세과세표준_계 + 항목(11)면세사업수입금액"))		
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End If    
	   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W12"), 15, 0)
		
	If Not ChkNotNull(lgcTB_17.GetData(2,"W13"), "수입금액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W13"), 15, 0)

	' -- 2006.03.21 개정 : 1. 조정후 수입금액명세서(A111)검증 추가 
	If  UNICDbl(lgcTB_17.GetData(2,"W13"), 0) <> dblAmtSumW4  Then
	    blnError = True
		Call SaveHTFError(lgsPGM_ID,lgcTB_17.GetData(2,"W13") & " <> " & dblAmtSumW4 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "1. 업종별 수입금액명세서의 수입금액_합계(코드 99)","2. 부가가치세 과세표준과 수입금액 차액검토의 (12) 수입금액"))		
	End If

		
	If  ChkNotNull(lgcTB_17.GetData(2,"W14"), "차액") Then 
	    '항목(14)차액 : 항목(12)합계 - 항목(13)수입금액 
	    IF UNICDbl(lgcTB_17.GetData(2,"W14"),0) <> UNICDbl(lgcTB_17.GetData(2,"W12"),0)  - UNICDbl(lgcTB_17.GetData(2,"W13"),0)  Then
	       Call SaveHTFError(lgsPGM_ID,UNICDbl(lgcTB_17.GetData(2,"W12"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "항목(14)차액","항목(12)합계 - 항목(13)수입금액"))		
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(2,"W14"), 15, 0)

	' -- 200603 개정		
	sHTFBody = sHTFBody & UNIChar("", 54) & vbCrLf	' -- 공란 


	' --제17호 조정후수입금액명세서 - 수입금액과의 차액내역 : 200603 개정 
	iSeqNo = 1	
	
	Do Until lgcTB_17.EOF(3)
		If UNICDbl(lgcTB_17.GetData(3,"W9"), 0) > 0 Then
	
			sHTFBody = sHTFBody & "85"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
			If Trim(lgcTB_17.GetData(3,"W8")) = "차액계" Then
				sHTFBody = sHTFBody & "999999"
			Else
				sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			End If
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W8"), "수입금액차액_구분") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(3,"W8"), 20)			
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W15"), "수입금액차액_코드") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(3,"W15"), 2, 0)	
		
			If Not ChkNotNull(lgcTB_17.GetData(3,"W15"), "수입금액차액_금액") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_17.GetData(3,"W9"), 15, 0)	
		
			'If Not ChkNotNull(lgcTB_17.GetData(3,"W_REMARK"), "수입금액차액_비고") Then blnError = True	' 2006.03.06 수정 
			sHTFBody = sHTFBody & UNIChar(lgcTB_17.GetData(3,"W_REMARK"), 20)			
		
			sHTFBody = sHTFBody & UNIChar("", 31) & vbCrLf	' -- 공란 
		
			iSeqNo = iSeqNo + 1
		End If
		
		lgcTB_17.MoveNext(3) 
		
	Loop
	

	' ----------- 
	PrintLog "WriteLine2File : 33" & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_17 = Nothing	' -- 메모리해제 
	
End Function


%>
