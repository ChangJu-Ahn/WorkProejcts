
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제16호 수입금액 조정명세서 
'*  3. Program ID           : W2105MA1
'*  4. Program Name         : W2105MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_16

Set lgcTB_16 = Nothing ' -- 초기화 

Class C_TB_16
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
		blnData1 = True : blnData2 = True : blnData3 = True
		
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
		Call SubMakeSQLStatements("D1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

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
	
	
	Function Clone(Byref pRs)
	   Set pRs = lgoRs1.clone   '복제 
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
	
	Function MoveFist(Byval pType)
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
				lgStrSQL = lgStrSQL & " FROM TB_16H	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_16D1	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				
		  Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_16D2	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W2105MA1
	Dim A111

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W2105MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
  ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W2105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W2105MA1"

	Set lgcTB_16 = New C_TB_16		' -- 해당서식 클래스 
	
	If Not lgcTB_16.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W2105MA1

	'==========================================
	' -- 제16호 수입금액 조정명세서 오류검증 
	' -- 1. 수입조정계산 
	iSeqNo = 1	
	
	Do Until lgcTB_16.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "SEQ_NO"), 6)
		End If
				
		If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(1, "W1_NM"), "계정 항목") Then blnError = True	
			If Not ChkNotNull(lgcTB_16.GetData(1, "W2_NM"), "계정 항목") Then blnError = True	
		End If
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "W1_NM"), 50)
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "W2_NM"), 50)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W3"), "결산서상 수입금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W3"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W4"), "조정가산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(1, "W5"), "조정차감") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W5"), 15, 0)
		
		If  ChkNotNull(lgcTB_16.GetData(1, "W6"), "조정후 수입금액") Then
		
			If unicdbl(lgcTB_16.GetData(1, "W6"),0) <> unicdbl(lgcTB_16.GetData(1, "W3"),0) + unicdbl(lgcTB_16.GetData(1, "W4"),0) - unicdbl(lgcTB_16.GetData(1, "W5"),0)  then
			    Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_16.GetData(1, "W6"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "조정후 수입금액","항목(3)결산서상수입금액 + 항목(4)조정_가산 - 항목(5)조정_차감"))
				 blnError = True	
			End if
			
			If UNICDbl(lgcTB_16.GetData(1, "SEQ_NO"), 0) = 999999 Then
			   Set cDataExists.A111  = new C_TB_17		' -- W2107MA1_HTF.asp 에 정의됨 
			 '  Call SubMakeSQLStatements_W2105MA1("A111",iKey1, iKey2, iKey3)   
			   cDataExists.A111.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			   cDataExists.A111.WHERE_SQL = ""	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			   
					If Not cDataExists.A111.LoadData() Then
					       blnError = True
						  Call SaveHTFError(lgsPGM_ID, "제17호 조정후 수입금액 명세서 ", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
					Else
				
					    
				        Call cDataExists.A111.Find ("1" ,"Code_no = '99'")   '코드가 99인것과 비교 
					   If unicdbl(cDataExists.A111.GetData(1,"W4"),0) <> unicdbl(lgcTB_16.GetData(1, "W6"),0) then
					      Call SaveHTFError(lgsPGM_ID,unicdbl(cDataExists.A111.getdata(1,"W4"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "조정후 수입금액","제17호 조정후 수입금액 명세서의 코드 99의 금액"))
					       blnError = True
					   End if
					   
					End if
				Set  cDataExists.A111 = nothing
			   
			End If
		Else
		    blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(1, "W6"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(1, "DESC1"), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- 공란 
		
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(1)	' -- 1번 레코드셋 
	Loop

	

	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	' -- 2. 작업진행률에 대한수입금액 
	iSeqNo = 1	: blnError = False
sHTFBody = ""
	Do Until lgcTB_16.EOF(2) 
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "SEQ_NO"), 40)
		End If
		
		
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then		 
		   If Not ChkNotNull(lgcTB_16.GetData(2, "W7"), "공사명") Then blnError = True	
		End if
		   
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "W7"), 50)
		If UNICDbl(lgcTB_16.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(2, "W8"), "도급자") Then blnError = True	
		End if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(2, "W8"), 30)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W9"), "도급금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W9"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W10"), "당해사업연도말 총공사비누적액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W11"), "총공사예정비") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W11"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W12"), "진행률") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W12"), 5, 2)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W13"), "익금산입액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W13"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W14"), "전기말수입계상액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W14"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(2, "W15"), "당기회사수입계상액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W15"), 15, 0)
		
		If  ChkNotNull(lgcTB_16.GetData(2, "W16"), "조정액") Then 
		   '항목(13)익금산입액 - 항목(14)전기말수입계상액 - 항목(15)당기회사수입계상액 +10,000 ~ -10,000
		   sTmp= unicdbl(lgcTB_16.GetData(2, "W13"),0) + unicdbl(lgcTB_16.GetData(2, "W14"),0) -unicdbl(lgcTB_16.GetData(2, "W15"),0) 
		   If unicdbl(lgcTB_16.GetData(2, "W16"),0) = sTmp + 10000 and unicdbl(lgcTB_16.GetData(2, "W16"),0) >= sTmp - 10000 then
		   Else
		       Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_16.GetData(2, "W16"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "조정액","항목(13)익금산입액 - 항목(14)전기말수입계상액 - 항목(15)당기회사수입계상액 +10,000 ~ -10,000"))
				blnError = True
		   End if
		   
		Else
		   blnError = True	
		End if
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(2, "W16"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 48) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(2)	' -- 1번 레코드셋 
	Loop
			

	PrintLog "Write2File : " & sHTFBody
 
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If


		' -- 3. 기타 수입금 
	iSeqNo = 1	: blnError = False
	sHTFBody = ""
	Do Until lgcTB_16.EOF(3) 
		sHTFBody = sHTFBody & "85"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_16.GetData(3, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		Else
			sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "SEQ_NO"), 40)
		End If
		'Response.End 
		
		'zzzz 합계는 체크 필요없음?
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "W17"), "구분") Then blnError = True
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "W17"), 50)
		
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "W18"), "근거법령") Then blnError = True	
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "W18"), 80)
		
		If Not ChkNotNull(lgcTB_16.GetData(3, "W19"), "수입금액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(3, "W19"), 15, 0)
		
		If Not ChkNotNull(lgcTB_16.GetData(3, "W20"), "대응원가") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_16.GetData(3, "W20"), 15, 0)
		
		If UNICDbl(lgcTB_16.GetData(3, "CHILD_SEQ_NO"), 0) <> 999999 Then
			If Not ChkNotNull(lgcTB_16.GetData(3, "DESC2"), "비고") Then blnError = True	
		end if	
		sHTFBody = sHTFBody & UNIChar(lgcTB_16.GetData(3, "DESC2"), 50)
		
		sHTFBody = sHTFBody & UNIChar("", 28) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_16.MoveNext(3)	' -- 1번 레코드셋 
	Loop

	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
			
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_16 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W2105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A111" '-- 외부 참조 SQL
           lgStrSQL = ""
      
	End Select
	PrintLog "SubMakeSQLStatements_W2105MA1 : " & lgStrSQL
End Sub
%>
