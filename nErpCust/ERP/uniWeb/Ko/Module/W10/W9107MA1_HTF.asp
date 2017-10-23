<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제52호 특수관계자간 거래명세서 
'*  3. Program ID           : W9107MA1
'*  4. Program Name         : W9107MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_52

Set lgcTB_52 = Nothing ' -- 초기화 

Class C_TB_52
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
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

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
	            lgStrSQL = lgStrSQL & " SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_52H	WITH (NOLOCK) " & vbCrLf	' 서식52호 
				lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				If WHERE_SQL = "" Then	' -- 외부호출외에 썸값을 리턴한다. 개정서식 200603반영: 2/15일자메일내용참조 
					lgStrSQL = lgStrSQL & "	AND (W3 > 0 OR W10 > 0) " & vbCrLf		' -- 계가 0보다 큰것만 반영 
					lgStrSQL = lgStrSQL & "	UNION ALL " & vbCrLf
					lgStrSQL = lgStrSQL & "	SELECT 999999 SEQ_NO, '', '', SUM(W3), SUM(W4), SUM(W5), SUM(W6), SUM(W7), SUM(W8), SUM(W9), SUM(W10), SUM(W11), SUM(W12), SUM(W13), SUM(W14), SUM(W15), SUM(W16), MAX(INSRT_USER_ID), MAX(INSRT_DT), MAX(UPDT_USER_ID), MAX(UPDT_DT) " & vbCrLf
					lgStrSQL = lgStrSQL & " FROM TB_52H	WITH (NOLOCK) " & vbCrLf	
					lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
					lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
					lgStrSQL = lgStrSQL & " HAVING SUM(W3) > 0 OR SUM(W10) > 0 " & vbCrLf	
				End If
				
	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_52H2	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9107MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9107MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9107MA1"

	Set lgcTB_52 = New C_TB_52		' -- 해당서식 클래스 
	
	If Not lgcTB_52.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9107MA1

	'==========================================
	' -- 제52호 특수관계자간 거래명세서 오류검증 
	' -- 1. 매출및매입거래등 
	iSeqNo = 1	
	
	Do Until lgcTB_52.EOF(1) 
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If UNICDbl(lgcTB_52.GetData(1, "SEQ_NO"), 0) = 999999 Then
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(1, "SEQ_NO"), 6)
		Else
			' -- 합계가 아닐때만 검증: 200603개정 
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		
			If Not ChkNotNull(lgcTB_52.GetData(1, "W1"), "법인명(성명)") Then blnError = True	
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(1, "W1"), 60)
	  
	  
			If  ChkNotNull(lgcTB_52.GetData(1, "W2"), "사업자등록번호(주민등록번호)") Then 
				If  Len(Replace(lgcTB_52.GetData(1, "W2"),"-","") ) <> 10  and Len(Replace(lgcTB_52.GetData(1, "W2"),"-","") ) <> 13 then 
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W2"), UNIGetMesg(TYPE_CHK_CHARNUM, "사업자등록번호(주민등록번호)","10 이거나 13"))
					blnError = True	
				End If
			   
			    If lgcCompanyInfo is Nothing Then
					blnError = True	
					Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("전자신고 Conversion프로그램에서 코드A100 법인기초 정보관리도 체크되어야 합니다.", "",""))
					Exit Function
			    End If
			    
				If replace(lgcCompanyInfo.OWN_RGST_NO,"-","") = Replace(lgcTB_52.GetData(1, "W2"),"-","") Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W2"), UNIGetMesg("납세자 사업자번호와 거래상대방 사업자번호는 같을 수 없습니다", "",""))
					blnError = True	
				 End If
		
			Else
			   blnError = True	
			End If

		End If
		
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_52.GetData(1, "W2")), 13)
		
		If  ChkNotNull(lgcTB_52.GetData(1, "W3"), "매출거래등_계") Then
		    '매출거래등_계 = 항목 (9)유형_재고자산 + (10)유형_기타 + (11)무형자산 + (12)용역+ (13)금전대부 + (14)기타 
		       sTmp =  UNICDbl(lgcTB_52.GetData(1, "W4"),0) +  UNICDbl(lgcTB_52.GetData(1, "W5"),0) + UNICDbl(lgcTB_52.GetData(1, "W6"),0) + _
		               UNICDbl(lgcTB_52.GetData(1, "W7"),0) +  UNICDbl(lgcTB_52.GetData(1, "W8"),0) +  UNICDbl(lgcTB_52.GetData(1, "W9"),0) 
				If  UNICDbl(lgcTB_52.GetData(1, "W3"),0) <> sTmp Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W3"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_52.GetData(1, "W1") & "의 매출거래등_계","유형_재고자산 + 유형_기타 + 무형자산 + 용역+ 금전대부 + 기타"))
				     blnError = True	
				End if
		
		    
		Else
			blnError = True	
		End If	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W3"), 15, 0)
				
		If Not ChkNotNull(lgcTB_52.GetData(1, "W4"), "매출거래등_유형자산 재고자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W4"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W5"), "매출거래등_유형자산 기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W5"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W6"), "매출거래등_무형자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W6"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W7"), "매출거래등_용역") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W7"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W8"), "매출거래등_금전대부") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W8"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W9"), "매출거래등_기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W9"), 15, 0)
		
		If  ChkNotNull(lgcTB_52.GetData(1, "W10"), "매입거래등_계") Then
		     '매입거래등_계 = - 항목 (16)유형_재고자산 + (17)유형_기타 + (18)무형자산 + (19)용역+ (20)금전대부 + (21)기타,0) + _
		       sTmp =  UNICDbl(lgcTB_52.GetData(1, "W11"),0) +  UNICDbl(lgcTB_52.GetData(1, "W12"),0) + UNICDbl(lgcTB_52.GetData(1, "W13"),0) + _
		               UNICDbl(lgcTB_52.GetData(1, "W14"),0) +  UNICDbl(lgcTB_52.GetData(1, "W15"),0) +  UNICDbl(lgcTB_52.GetData(1, "W16"),0) 
				If  UNICDbl(lgcTB_52.GetData(1, "W10"),0) <> sTmp Then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(1, "W10"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_52.GetData(1, "W1") & "의 매입거래등_계","유형_재고자산 + 유형_기타 + 무형자산 + 용역+ 금전대부 + 기타"))
				     blnError = True	
				End if
		Else
			blnError = True	
		End if	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W10"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W11"), "매입거래등_유형자산 재고자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W11"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W12"), "매입거래등_유형자산 기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W12"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W13"), "매입거래등_무형자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W13"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W14"), "매입거래등_용역") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W14"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W15"), "매입거래등_금전대부") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W15"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(1, "W16"), "매입거래등_기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W16"), 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 5) & vbCrLf	' -- 공란 

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_52.MoveNext(1)	' -- 1번 레코드셋 
	Loop

	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = "" ' -- 자본거래는 뒤에 출력해야되므로 꼭 초기화: 200603
	' -- 2. 자본거래 
	iSeqNo = 1	
	
	Do Until lgcTB_52.EOF(2) 
		sHTFBody = sHTFBody & "83"
		'sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		sHTFBody = sHTFBody & UNIChar("A232", 4)		' -- 200603 개정서식 
		
		'If UNICDbl(lgcTB_52.GetData(2, "SEQ_NO"), 0) <> 999999 Then
			sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		'Else
		'	sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "SEQ_NO"), 6)
		'End If
				
		If Not ChkNotNull(lgcTB_52.GetData(2, "W17"), "법인명(상호 또는 성명)") Then blnError = True	
		sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W17"), 60)
		
	
		
		If  ChkNotNull(lgcTB_52.GetData(2, "W18"), "사업자등록번호(주민등록번호)") Then 
			If  Len(Replace(lgcTB_52.GetData(2, "W18"),"-","") ) <> 10  and Len(Replace(lgcTB_52.GetData(2, "W18"),"-","") ) <> 13 then 
			    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(2, "W18"), UNIGetMesg(TYPE_CHK_CHARNUM, "사업자등록번호(주민등록번호)","10 이거나 13"))
				blnError = True	
			End If
		   
			If replace(lgcCompanyInfo.OWN_RGST_NO,"-","") = Replace(lgcTB_52.GetData(2, "W18"),"-","") Then
			    Call SaveHTFError(lgsPGM_ID, lgcTB_52.GetData(2, "W18"), UNIGetMesg("납세자 사업자번호와 거래상대방 사업자번호는 같을 수 없습니다", "",""))
				blnError = True	
			 End If
		
		Else
		   blnError = True	
		End If
		 
		 
		 sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_52.GetData(2, "W18")), 13)
		
		'If Not ChkNotNull(lgcTB_52.GetData(2, "W19"), "증자,감자 구분") Then blnError = True
		If lgcTB_52.GetData(2, "W19") = "0" Then
			sHTFBody = sHTFBody & UNIChar("", 1)
		Else
			if Not ChkBoundary("1,2",lgcTB_52.GetData(2, "W19"),"증자,감자 구분") then   blnError = True	
						
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W19"), 1)
			
			If Not ChkNotNull(lgcTB_52.GetData(2, "W21"), "증자,감자 일자") Then blnError = True	'증감자 체크시 일자는 필수)
		End If
			
		sHTFBody = sHTFBody & UNI8Date(lgcTB_52.GetData(2, "W21"))
	
		If Not ChkNotNull(lgcTB_52.GetData(2, "W22"), "증자(또는 감자)전_액면총액 ") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W22"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W23"), "증자(또는 감자)전_지분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W23"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W24"), "증자(또는 감자)후_액면총액") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W24"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W25"), "증자(또는 감자)후_지분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W25"), 5, 2)
		
		'If Not ChkNotNull(lgcTB_52.GetData(2, "W26"), "합병,분할합병 구분") Then blnError = True
		If lgcTB_52.GetData(2, "W26") = "0" Then
			sHTFBody = sHTFBody & UNIChar("", 1)
		Else
		   	if Not ChkBoundary("1,2",lgcTB_52.GetData(2, "W26"),"합병,분할합병 구분") then   blnError = True	
						
			sHTFBody = sHTFBody & UNIChar(lgcTB_52.GetData(2, "W26"), 1)
			  
			If Not ChkNotNull(lgcTB_52.GetData(2, "W28"), "합병,분할합병 일자") Then blnError = True	'합병,분할합병 구분체크시 일자는 필수)
		End If	
		
		If lgcTB_52.GetData(2, "W19") = "0" And lgcTB_52.GetData(2, "W26") = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", UNIGetMesg("증자/감자, 합병/분할합병 중 적어도 1군데는 체크되어야 됩니다.", "",""))
		End If

		
		sHTFBody = sHTFBody & UNI8Date(lgcTB_52.GetData(2, "W28"))

		If Not ChkNotNull(lgcTB_52.GetData(2, "W29_1"), "합병법인등 순자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W29_1"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W29_2"), "합병법인등 지분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W29_2"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W30_1"), "피합병법인등 순자산") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W30_1"), 15, 0)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W30_2"), "피합병법인등 지분") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W30_2"), 5, 2)
		
		If Not ChkNotNull(lgcTB_52.GetData(2, "W31"), "합병비율") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_52.GetData(1, "W31"), 5, 2)
		
		sHTFBody = sHTFBody & UNIChar("", 12) & vbCrLf	' -- 공란 
	
		iSeqNo = iSeqNo + 1
		
		Call lgcTB_52.MoveNext(2)	' -- 2번 레코드셋 
	Loop
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		'Call Write2File(sHTFBody)	' -- 200603 개정 
		Call PushRememberDoc(sHTFBody)	' -- 바로 출력하지 않고 기억시킨다(inc_HomeTaxFunc.asp에 정의)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_52 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL
			
			lgStrSQL = ""
			' -- 표준손익계산서(A115,A116)의 당기순손익(일반법인은 코드(82) 금융.보험.증권업법인은 코드(73))과 일치하지 않으면 오류 
			lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- 표준손익계산서 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- 법인구분(일반/금융)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '82'"		 	 & vbCrLf	' -- 법인구분(일반)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '73'"		 	 & vbCrLf	' -- 법인구분(금융)
			End If
	End Select
	PrintLog "SubMakeSQLStatements_W9107MA1 : " & lgStrSQL
End Sub
%>
