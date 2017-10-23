<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제64호 감가상각방법신고서 
'*  3. Program ID           : W9115MA1
'*  4. Program Name         : W9115MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_63

Set lgcTB_63 = Nothing ' -- 초기화 

Class C_TB_63
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
		Call SubMakeSQLStatements("A",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("B",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

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
			Case 2
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
				lgStrSQL = lgStrSQL & " FROM TB_63H	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "A"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_63A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
			
			Case "B"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
	            
	            If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보의 사업개시일(창립일) 불러온다 
					lgStrSQL = lgStrSQL & " , B.FOUNDATION_DT " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " FROM TB_63B	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				
				If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보를 조인한다.
					lgStrSQL = lgStrSQL & " INNER JOIN TB_COMPANY_HISTORY B WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9115MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9115MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9115MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9115MA1"

	Set lgcTB_63 = New C_TB_63		' -- 해당서식 클래스 
	
	If Not lgcTB_63.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9115MA1

	'==========================================
	' -- 제64호 감가상각방법신고서 오류검증 
	' -- 1. 매출및매입거래등 
	iSeqNo = 1	
	
	If lgcTB_63.EOF(2) Then
			' -- 그리드가 존재하지않느다면 빈거 1개 생성 
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
			
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
					
			sHTFBody = sHTFBody & UNIChar("", 30)
			
			sHTFBody = sHTFBody & UNIChar("", 2)
			sHTFBody = sHTFBody & UNIChar("", 2)
			
			sHTFBody = sHTFBody & UNINumeric(0, 2, 0)
			sHTFBody = sHTFBody & UNINumeric(0, 2, 0)
						
			sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- 공란 
	Else
	
		Do Until lgcTB_63.EOF(2) 
			sHTFBody = sHTFBody & "83"
			sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
			
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
			'	내용연수범위, 신고내용연수, 변경내용연수 모두 ‘0’이면 검증 제외	
			If Not ChkNotNull(lgcTB_63.GetData(2, "W8"), "자산 및 업종명") Then 
			
			    if UNINumeric(lgcTB_63.GetData(2, "W9_Fr"), 2, 0)  <> 0 Or UNINumeric(lgcTB_63.GetData(2, "W9_To"), 2, 0) <> 0 Or  UNINumeric(lgcTB_63.GetData(2, "W10"), 2, 0)  <> 0 Or UNINumeric(lgcTB_63.GetData(2, "W11"), 2, 0) <> 0 then
			       blnError = True	
			    End if   
			End if    
			sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(2, "W8"), 30)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W9_Fr"), "내용연수범위_From") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W9_Fr"), 2, 0)
					
			If Not ChkNotNull(lgcTB_63.GetData(2, "W9_To"), "내용연수범위_To") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W9_To"), 2, 0)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W10"), "신고내용연수") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W10"), 2, 0)
			
			If Not ChkNotNull(lgcTB_63.GetData(2, "W11"), "변경내용연수") Then blnError = True	
			sHTFBody = sHTFBody & UNINumeric(lgcTB_63.GetData(2, "W11"), 2, 0)
			
			
			sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(2, "W12"), 50)
			
			sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- 공란 

			iSeqNo = iSeqNo + 1
			
			Call lgcTB_63.MoveNext(2)	' -- 1번 레코드셋 
		Loop
	End If
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	' -- 2. 자본거래 
	
	Do Until lgcTB_63.EOF(3) 
	
		sHTFBody = sHTFBody & "84"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

		If  ChkNotNull(lgcTB_63.GetData(3, "FOUNDATION_DT"), "사업개시일") Then 
		    if ChkDate(lgcTB_63.GetData(3, "FOUNDATION_DT") , "사업개시일") = False  Then blnError = True	
		Else
		    blnError = True	
		End if    
		sHTFBody = sHTFBody & UNI8Date(lgcTB_63.GetData(3, "FOUNDATION_DT"))
		
		If Not ChkNotNull(lgcTB_63.GetData(3, "W7"), "변경방법적용사업연도") Then blnError = True	
		sHTFBody = sHTFBody & UNI8Date(lgcTB_63.GetData(3, "W7"))
		
		If Trim(lgcTB_63.GetData(3, "W13_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W13_A")) <> null  Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W13_A"), "유형고정자산_신고상각방법") Then blnError = True
		Else
		 blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_A"), 1)
	
	
		
		If  Trim(lgcTB_63.GetData(3, "W13_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W13_B")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W13_B"), "유형고정자산_변경상각방법") Then blnError = True
		Else
		 blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_B"), 1)
		
		
		'유형고정자산_변경사유 
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W13_C"), 50)
		
		
	
		If  Trim(lgcTB_63.GetData(3, "W14_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W14_A")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W14_A"), "광업권_신고상각방법법") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_A"), 1)
		
	
		If  Trim(lgcTB_63.GetData(3, "W14_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W14_B")) <> null    Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W14_B"), "광업권_변경상각방법") Then blnError = True
		Else
		   blnError = True	
		End if 
		   sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_B"), 1)
		

		'광업권_변경사유 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W14_C"), 50)
		
	
		If Trim(lgcTB_63.GetData(3, "W15_A")) <> "" Or  Trim(lgcTB_63.GetData(3, "W15_A")) <> null   Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W15_A"), "광업용고정자산_신고상각방법") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(3, "W15_A"), 1)
		

		If  Trim(lgcTB_63.GetData(3, "W15_B")) <> "" Or  Trim(lgcTB_63.GetData(3, "W15_B")) <> null Then
		   If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_63.GetData(3, "W15_B"), "광업용고정자산_변경상각방법") Then blnError = True
		Else
		   blnError = True	
		End if 
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(1, "W15_B"), 1)

        '광업용고정자산_변경사유 
	
		sHTFBody = sHTFBody & UNIChar(lgcTB_63.GetData(1, "W15_C"), 50)
	
		if Trim(lgcTB_63.GetData(3, "W13_A")) = "" and  Trim(lgcTB_63.GetData(3, "W14_A")) = "" and Trim(lgcTB_63.GetData(3, "W15_A")) = ""  then
		   Call SaveHTFError(lgsPGM_ID, Trim(lgcTB_63.GetData(3, "W13_A")), UNIGetMesg(TYPE_CHK_NULL, " 항목(13) 유형고정자산_신고상각방법, 항목(14) 광업권_신고상각방법,항목(15) 광업용고정자산_신고상각방법 중 하나",""))
		   blnError = True	
		end if
 	
		sHTFBody = sHTFBody & UNIChar("", 22) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_63.MoveNext(3)	' -- 2번 레코드셋 
	Loop
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_63 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9115MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9115MA1 : " & lgStrSQL
End Sub
%>
