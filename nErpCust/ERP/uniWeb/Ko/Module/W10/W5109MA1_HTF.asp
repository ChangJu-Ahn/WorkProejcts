<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제22호 기부금명세서 
'*  3. Program ID           : W5109MA1
'*  4. Program Name         : W5109MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_22

Set lgcTB_22 = Nothing ' -- 초기화 

Class C_TB_22
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
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

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
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_22H	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_22D	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W5109MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W5109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W5109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W5109MA1"

	Set lgcTB_22 = New C_TB_22		' -- 해당서식 클래스 
	
	If Not lgcTB_22.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W5109MA1

	'==========================================
	' -- 제22호 기부금명세서 오류검증 
	' -- 1. 매출및매입거래등 
	'sHTFBody = "83"
	'sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	sHTFBody =""
	iSeqNo = 1	
	
	Do Until lgcTB_22.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		
		If Not ChkNotNull(lgcTB_22.GetData(1, "W2"), lgcTB_22.GetData(1, "W1") & "구분코드") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W2"), 2)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W3"), 40)
		
		If lgcTB_22.GetData(1, "W2") <> "20" And lgcTB_22.GetData(1, "W2") <> "50" Then	
		    '날짜 형식체크 
			If Not (ChkNotNull(lgcTB_22.GetData(1, "W4"), lgcTB_22.GetData(1, "W1") & "연월") and  ChkDate(lgcTB_22.GetData(1, "W4"),lgcTB_22.GetData(1, "W1") & "연월"))  Then blnError = True	
			 
			'기부일자가 사업년도 기간 이내가 아니면 오류 
			If Not (UNI6Date(lgcTB_22.GetData(1, "W4"))  >= UNI6Date(lgcCompanyInfo.FISC_START_DT) and UNI6Date(lgcTB_22.GetData(1, "W4"))  <=UNI6Date(lgcCompanyInfo.FISC_END_DT))   Then 
			  Call SaveHTFError(lgsPGM_ID,lgcTB_22.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_22.GetData(1, "W1") & "기부일자", "사업년도기간"))
			  blnError = True	
			End if  
			
		End If
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_22.GetData(1, "W4")), 6)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W5"), 80)
		
		
		If lgcTB_22.GetData(1, "W2") <> "20" Then	
			If Not ChkNotNull(lgcTB_22.GetData(1, "W6"), "기부처") Then blnError = True	
		End If
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W6"), 60)
		
		
	   if  	UNIRemoveDash(lgcTB_22.GetData(1, "W7")) <> "" then
			If  Len(UNIRemoveDash(lgcTB_22.GetData(1, "W7"))) <> 10  and Len(UNIRemoveDash(lgcTB_22.GetData(1, "W7")) ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W7"), UNIGetMesg(TYPE_CHK_CHARNUM, lgcTB_22.GetData(1, "W6") & "사업자등록번호(주민등록번호)","10 이거나 13"))
				blnError = True	
			End If
			   
			If UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO) = UNIRemoveDash(lgcTB_22.GetData(1, "W7"))  Then
			    Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W7"), UNIGetMesg("납세자 사업자번호와 거래상대방 사업자번호는 같을 수 없습니다", "",""))
				blnError = True	
			 End If
		End if	 
		 
		 sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_22.GetData(1, "W7")), 13)
		
		If  ChkNotNull(lgcTB_22.GetData(1, "W8"), "금액") Then 
		    If Unicdbl(lgcTB_22.GetData(1, "W8"),0) < 0 then 
		         Call SaveHTFError(lgsPGM_ID, lgcTB_22.GetData(1, "W8"), UNIGetMesg(TYPE_POSITIVE, "금액",""))
				blnError = True	
		    End if
		Else
			blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_22.GetData(1, "W8"), 15, 0)
	 
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(1, "W_DESC"), 80)
		
		sHTFBody = sHTFBody & UNIChar("", 42) & vbCrLf	' -- 공란 
		
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_22.MoveNext(1)	' -- 1번 레코드셋 
	Loop
	
	
	         '999990 : 가. 법 제24조제2항 기부금(법정기부금, 코드 10)의 총합 
             '999991 : 나. 조세특례제한법 제76조 기부금(정치자금, 코드 20)의 총합 
             '999992 : 다. 조세특례제한법 제73조 제1항 제1호 기부금(코드 60)의 총합 
             '999993 : 라. 조세특례제한법 제73조 제1항 제2호 내지 제15호 기부금(코드 30)의 총합 
             '999994 : 마. 법 제24조 제1항 기부금(지정기부금, 코드 40)의 총합 
             '999995 : 바. 조세특례제한법 제73조 제2항 기부금(코드 70)의 총합 
             '999996 : 마. 기타 기부금(코드 50)의 총합 
             '999999 : 합계로서 999990 ~ 999996 의 총합을 말합니다.

	
	' -- 2. 자본거래 
	iSeqNo = 999990	
	
	Do Until lgcTB_22.EOF(2) 
		If lgcTB_22.GetData(2, "W9_CD") <> "20" Then	' -- 20 정치가 제거됨 
	
			if   UNIChar(lgcTB_22.GetData(2, "W9_CD"), 2)  = "99" Then 
				iSeqNo = 999999
			End if
	
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			 
			If Not ChkNotNull(lgcTB_22.GetData(2, "W9_CD"), "구분코드") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(2, "W9_CD"), 2)
		
			sHTFBody = sHTFBody & UNIChar("", 40)
			sHTFBody = sHTFBody & UNIChar("", 6)
			sHTFBody = sHTFBody & UNIChar("", 80)
			sHTFBody = sHTFBody & UNIChar("", 60)
			sHTFBody = sHTFBody & UNIChar("", 13)
		
			If Not ChkNotNull(lgcTB_22.GetData(2, "W9_AMT"), "금액") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_22.GetData(2, "W9_AMT"), 15, 0)
		
			sHTFBody = sHTFBody & UNIChar(lgcTB_22.GetData(2, "W9_DESC"), 80)

			sHTFBody = sHTFBody & UNIChar("", 42) & vbCrLf	' -- 공란 
		End If
				
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_22.MoveNext(2)	' -- 1번 레코드셋 
	Loop
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_22 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W5109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W5109MA1 : " & lgStrSQL
End Sub
%>
