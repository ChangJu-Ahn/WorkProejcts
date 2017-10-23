<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제3호 이익잉여금처분(결손금처리)
'*  3. Program ID           : W1119MA1
'*  4. Program Name         : W1119MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_3_3_4

Set lgcTB_3_3_4 = Nothing ' -- 초기화 

Class C_TB_3_3_4
	' -- 테이블의 컬럼변수 
	Dim W_TYPE
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W8
	Dim W10
	Dim W11
	Dim W12
	Dim W13
	Dim W14
	Dim W15
	Dim W16
	Dim W17
	Dim W18
	Dim W19
	Dim W20
	Dim W25
	
	' -- 개정 2006.03
	Dim W26	
	Dim W27
	Dim W28
	' ---
	
	Dim W30
	Dim W31
	Dim W32
	Dim W33
	Dim W34
	Dim W35
	Dim W40
	Dim W41
	Dim W42
	Dim W43
	Dim W44
	Dim W50
	Dim W_DT
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
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

		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 
		
			If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
				If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				    IF CHK_COMPANY = TRUE THEN
					   Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
					 END IF	  
				End If
			    Exit Function
			End If
		

		' 멀티행이지만 첫행을 리턴 
		Call GetData
		
		Call CloseRs
		
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
	
	Function MoveFirst()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W_TYPE		= lgoRs1("W_TYPE")
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			W8			= lgoRs1("W8")
			W10			= lgoRs1("W10")
			W11			= lgoRs1("W11")
			W12			= lgoRs1("W12")
			W13			= lgoRs1("W13")
			W14			= lgoRs1("W14")
			W15			= lgoRs1("W15")
			W16			= lgoRs1("W16")
			W17			= lgoRs1("W17")
			W18			= lgoRs1("W18")
			W19			= lgoRs1("W19")
			W20			= lgoRs1("W20")
			W25			= lgoRs1("W25")
			
			' -- 개정 2006.03
			W26			= lgoRs1("W26")
			W27			= lgoRs1("W27")
			W28			= lgoRs1("W28")
			
			
			W30			= lgoRs1("W30")
			W31			= lgoRs1("W31")
			W32			= lgoRs1("W32")
			W33			= lgoRs1("W33")
			W34			= lgoRs1("W34")
			W35			= lgoRs1("W35")
			W40			= lgoRs1("W40")
			W41			= lgoRs1("W41")
			W42			= lgoRs1("W42")
			W43			= lgoRs1("W43")
			W44			= lgoRs1("W44")
			W50			= lgoRs1("W50")
			W_DT		= lgoRs1("W_DT")
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
				lgStrSQL = lgStrSQL & " FROM TB_3_3_4 A WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W1119MA1
	Dim A100_BASIC
	Dim A100
	Dim A113
	Dim A115
	Dim A110
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W1119MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, strMAIN_IND, dblCode5678 ,dblCode8273
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1119MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1119MA1"

	Set lgcTB_3_3_4 = New C_TB_3_3_4		' -- 해당서식 클래스 
	
	If Not lgcTB_3_3_4.LoadData Then Exit Function			
    Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
	
	'==========================================
	' -- 제3호 이익잉여금처분(결손금처리) 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	If Not ChkNotNull(lgcTB_3_3_4.W_DT, "처분확정일") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_3_3_4.W_DT)


	If lgcCompanyInfo.Comp_type2 = "2" then   '금융법인인경우 
	   '코드(10)합계가 음수이고, 코드(15)배당금이 ‘0’이 아니면I.처분전결손금 항목들의 값을 I.처분전이익잉여금 항목들에 입력해야 함.
	  ' 그 외에는 코드(10)합계가 음수인 경우I.처분전이익잉여금 항목들의 값을 I.처분전결손금 항목들에 입력해야 함.
    
		 If unicdbl(lgcTB_3_3_4.W10,0) < 0  and  unicdbl(lgcTB_3_3_4.W15,0) <> 0 Then
			Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg("금융법인인경우 합계(10)이 음수이고 배당금이(15)가 0이아니면 1.처분전 결손금 항목값들의 값을 처분전이익잉여금 항목들에 입력해야함", "",""))
			blnError = True	
		  END IF
	Else	  	
	     if unicdbl(lgcTB_3_3_4.W10,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg("처분전이익잉여금 항목들의 값을 처분전 결손금 항목들에 입력해야함", "",""))
			blnError = True	
		  END IF
	End If



	'* 코드 02 + 03 + 04 - 05 + 06
	'*표준대차대조표의 처분전이익잉여금 또는 처리전결손금과 일치  
	 ' 일반법인 : 표준대차대조표(일반법인)(A113) 의 코드 (56)               처분전이익잉여금또는처리전결손금 
	 ' 금융법인 : 표준대차대조표(금융법인)(A114) 의 코드 (78)              처분전이익잉여금또는처리전결손금 


	If  ChkNotNull(lgcTB_3_3_4.W1, "처분전이익잉여금") Then 

		if lgcCompanyInfo.Comp_type2 = "1" then 
		    dblCode5678 =  Getdata_TB_3_3_4_A142("A113_1")
		   
		Else
		    dblCode5678=  Getdata_TB_3_3_4_A142("A114_1")
		End if   

		if unicdbl(dblCode5678,0) <> unicdbl(lgcTB_3_3_4.W1,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "처분전이익잉여금","표준대차대조표의 처분전이익잉여금 또는 처리전결손금"))
		    blnError = True	
		End if
		
		if unicdbl( lgcTB_3_3_4.W1,0)  <> unicdbl( lgcTB_3_3_4.W2,0) + unicdbl( lgcTB_3_3_4.W3,0) +  unicdbl( lgcTB_3_3_4.W4,0) -  unicdbl( lgcTB_3_3_4.W5,0) +  unicdbl( lgcTB_3_3_4.W6,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "처분전이익잉여금","코드 02 + 03 + 04 - 05 + 06"))
		    blnError = True	
		End if
		
		
		
		
			
			
	Else
	        blnError = True	
	End if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W1, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W2, "전기이월이익잉여금(또는 전기이월결손금)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W2, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W3, "회계변경의 누적효과") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W3, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W4, "전기오류수정이익(또는 전기오류수정손실)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W4, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W5, "중간배당액") Then 
	    if unicdbl(lgcTB_3_3_4.W5,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W5, UNIGetMesg("중간배당액 값이 음수입니다.", "",""))
	         blnError = True	
	    End if
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W5, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W6, "당기순이익(또는 당기순손실)") Then 
	  ' - 일반법인 : 표준손익계산서(일반법인)(A115)의 코드(82)당기순이익(순손실)
	  '- 금융법인 : 표준손익계산서(금융법인)(A116)의 코드(73)당기순이익(순손실)
	   if lgcCompanyInfo.Comp_type2 = "1" then 
		    dblCode8273 =  Getdata_TB_3_3_A142("A115_1")
		Else
		    dblCode8273=  Getdata_TB_3_3_A142("A116_1")
		End if 
		
		if unicdbl(dblCode8273,0) <> unicdbl(lgcTB_3_3_4.W6,0) then
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W6, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당기순이익","표준대차대조표의 당기순이익(순손실)"))
		    blnError = True	
		End if
	
	Else
		blnError = True	
    End if		
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W6, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W8, "임의적립금 등의 이입액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W8, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W10, "합계") Then
	    '- 코드(01)처분전이익잉여금 + 코드(08)임의적립금 등의 이입액 
	     if unicdbl( lgcTB_3_3_4.W10,0)  <> unicdbl( lgcTB_3_3_4.W1,0) + unicdbl( lgcTB_3_3_4.W8,0)  then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W10, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "합계","처분전이익잉여금(01) + 임의적립금 등의 이입액(08)"))
		    blnError = True	
		 End if
	Else
	     blnError = True	
	End if	 
		
		
		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W10, 15, 0)


	If  ChkNotNull(lgcTB_3_3_4.W11, "이익잉여금 처분액") Then
	     '- 코드 12 + 13 + 14 + 15 + 18 + 19 + 20 + 26 + 27 + 28 (2006.03개정)
	     if unicdbl( lgcTB_3_3_4.W11,0)  <> unicdbl( lgcTB_3_3_4.W12,0) + unicdbl( lgcTB_3_3_4.W13,0) + unicdbl( lgcTB_3_3_4.W14,0) + unicdbl( lgcTB_3_3_4.W15,0) + unicdbl( lgcTB_3_3_4.W18,0) + unicdbl( lgcTB_3_3_4.W19,0) + unicdbl( lgcTB_3_3_4.W20,0) + unicdbl( lgcTB_3_3_4.W26,0) + unicdbl( lgcTB_3_3_4.W27,0) + unicdbl( lgcTB_3_3_4.W28,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W11, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "이익잉여금 처분액","처분전이익잉여금(01) + 임의적립금 등의 이입액(08)"))
		    blnError = True	
		 End if
	Else
	     blnError = True	
	End if	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W11, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W12, "이익준비금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W12, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W13, "기타법정적립금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W13, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W14, "주식할인발행차금상각액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W14, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W15, "배당금") Then 
	    '코드 16 + 17
	    if unicdbl( lgcTB_3_3_4.W15,0)  <> unicdbl( lgcTB_3_3_4.w16,0) + unicdbl( lgcTB_3_3_4.w17,0)  then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W15, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "배당금","가.현금배당(16) + 나.주식배당(17)"))
		    blnError = True	
		 End if
	Else
	    blnError = True	
	End if
	
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W15, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W16, "현금배당") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W16, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W17, "주식배당") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W17, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W18, "사업확장적립금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W18, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W19, "감채적립금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W19, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W20, "기타적립금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W20, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W25, "차기이월이익잉여금") Then 
	     '코드 01 + 08 + 11
	    if unicdbl( lgcTB_3_3_4.W25,0)  <> unicdbl( lgcTB_3_3_4.w1,0) + unicdbl( lgcTB_3_3_4.w8,0) -unicdbl( lgcTB_3_3_4.w11,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W25, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차기이월이익잉여금","처분전이익잉여금(01) + 임의적립금 등의 이입액(08) - 이익잉여금 처분액(11)"))
		    blnError = True	
		 End if
	 
	Else
	    blnError = True	
	End if  
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W25, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W30, "처리전결손금") Then
	    If unicdbl(lgcTB_3_3_4.W30,0) <> 0 Then
				If unicdbl(lgcTB_3_3_4.W30,0) < 0 then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg("처리전결손금 값이 음수입니다.", "",""))
				     blnError = True	
				End if
	    
	    
				'코드 (31 + 32 + 33 - 34 - 35) X (-1)
				If(unicdbl( lgcTB_3_3_4.w31,0) + unicdbl( lgcTB_3_3_4.w32,0) + unicdbl( lgcTB_3_3_4.w33,0)-unicdbl( lgcTB_3_3_4.w34,0)-unicdbl( lgcTB_3_3_4.w35,0)) * -1 < 0 and unicdbl(lgcTB_3_3_4.W30,0) <> 0 then
				    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg("", "코드 (31 + 32 + 33 - 34 - 35) X (-1)의 결과가 음수이므로 이익잉여금처분계산서에 입력하는 경우이거나 코드(31,35)값에 표현이 오류입니다"))
				    blnError = True	
				 End If
				 
				 
				if lgcCompanyInfo.Comp_type2 = "1" then 
				    dblCode5678 =  Getdata_TB_3_3_4_A142("A113_1")
				   
				Else
				    dblCode5678=  Getdata_TB_3_3_4_A142("A114_1")
				End if   

				if abs(unicdbl(dblCode5678,0)) <> unicdbl(lgcTB_3_3_4.W30,0) then
				   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W30, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "처리전결손금","표준대차대조표의 처분전이익잉여금 또는 처리전결손금"))
				    blnError = True	
				End if
		 End If		
	
	    
	Else
		 blnError = True	
	End If	 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W30, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W31, "전기이월이익잉여금(또는 전기이월결손금)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W31, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W32, "회계변경의 누적효과") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W32, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W33, "전기오류수정이익(또는 전기오류수정손실)") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W33, 15, 0)


	
	If  ChkNotNull(lgcTB_3_3_4.W34, "중간배당액") Then 
	    if unicdbl(lgcTB_3_3_4.W34,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W34, UNIGetMesg("중간배당액 값이 음수입니다.", "",""))
	         blnError = True	
	    End if
	Else
	    blnError = True	
	End if   
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W34, 15, 0)
	

	If  ChkNotNull(lgcTB_3_3_4.W35, "당기순손실(또는 당기순이익)") Then 
	    ' - 일반법인 : 표준손익계산서(일반법인)(A115)의 코드(82)당기순이익(순손실)
	  '- 금융법인 : 표준손익계산서(금융법인)(A116)의 코드(73)당기순이익(순손실)
	  IF unicdbl(lgcTB_3_3_4.W35,0) <> 0 Then
			if lgcCompanyInfo.Comp_type2 = "1" then 
				    dblCode8273 =  Getdata_TB_3_3_A142("A115_1")
				Else
				    dblCode8273=  Getdata_TB_3_3_A142("A116_1")
				End if 
		
				if unicdbl(dblCode8273,0)*(-1) <> unicdbl(lgcTB_3_3_4.W35,0)  then
				   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W35, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당기순손실(또는 당기순이익","표준대차대조표의 당기순이익(순손실)"))
				    blnError = True	
				End if
		End If		
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W35, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W40, "결손금처리액") Then 
	
	    If unicdbl(lgcTB_3_3_4.W40,0) < 0 then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg("결손금처리액 값이 음수입니다.", "",""))
	         blnError = True	
	    End if
	    
	    
	     '-41 + 42 + 43 + 44
	     if unicdbl( lgcTB_3_3_4.W40,0)  <> unicdbl( lgcTB_3_3_4.W41,0) + unicdbl( lgcTB_3_3_4.W42,0) + unicdbl( lgcTB_3_3_4.W43,0) + unicdbl( lgcTB_3_3_4.W44,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "결손금처리액","코드(41 + 42 + 43 + 44)"))
		    blnError = True	
		 End if
	Else
		blnError = True	
	End if	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W40, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W41, "임의적립금이입액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W41, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W42, "기타법정적립금이입액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W42, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W43, "이익준비금이입액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W43, 15, 0)

	If Not ChkNotNull(lgcTB_3_3_4.W44, "자본잉여금이입액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W44, 15, 0)

	If  ChkNotNull(lgcTB_3_3_4.W50, "차기이월결손금") Then 
	   '30-40
	     if unicdbl( lgcTB_3_3_4.W50,0)  <>  unicdbl( lgcTB_3_3_4.W30,0) - unicdbl( lgcTB_3_3_4.W40,0) then
		    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W50, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차기이월결손금","코드(30-40)"))
		    blnError = True	
		 End if
	Else
	    blnError = True	
	End If 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W50, 15, 0)
	
	
	
	if unicdbl(lgcTB_3_3_4.W25,0) <> 0 and unicdbl( lgcTB_3_3_4.W50,0) <> 0 then
	   Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_4.W40, UNIGetMesg("차기이월이익잉여금(25)와 차기이월 결손금(50) 모두 금액을 입력할 수 없습니다","", ""))
		    blnError = True	
	End if

	' -- 2006.03 개정 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W26, 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W27, 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_4.W28, 15, 0)

	' -- 47호 갑 서식의 코드98 금액과 일치 비교 필요함 : 2006.03 
	If lgcTB_3_3_4.W_TYPE = "1" Then
		If UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0) > 0 And lgcCompanyInfo.Comp_type2 <> "2" Then	' 2006.03.24 개정서식 법인구분이 2가 아닐때 
			dblCode5678 =  Getdata_TB_47_A110("A110")

			if dblCode5678 <> UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0)  then
			   Call SaveHTFError(lgsPGM_ID, UNICdbl(lgcTB_3_3_4.W5, 0) + UNICdbl(lgcTB_3_3_4.W15, 0) + UNICdbl(lgcTB_3_3_4.W26, 0)  & " <> " &dblCode5678 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(05)중간배당액 + 코드(15)배당금 + 코드(26)이익처분에의한상여금", "제47호 주요계정명세서(갑)(A110)의 코드(98) 이익처분금액"))
			   blnError = True	
			End if

		End If
	Else
		If UNICdbl(lgcTB_3_3_4.W34, 0) > 0  And lgcCompanyInfo.Comp_type2 <> "2" Then	' 2006.03.24 개정서식 법인구분이 2가 아닐때 
			dblCode5678 =  Getdata_TB_47_A110("A110")

			if dblCode5678 <> UNICdbl(lgcTB_3_3_4.W34, 0)  then
			   Call SaveHTFError(lgsPGM_ID, UNICdbl(lgcTB_3_3_4.W34, 0) & " <> " &dblCode5678 , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(34)중간배당액", "제47호 주요계정명세서(갑)(A110)의 코드(98) 이익처분금액"))
			   blnError = True	
			End if
		End If
	
	End If	


	sHTFBody = sHTFBody & UNIChar("", 26)	' -- 공란 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_3_3_4 = Nothing	' -- 메모리해제 
	
End Function





Function CHK_COMPANY()
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , chkData1,chkData2,chkData3, chkData4
 CHK_COMPANY = FALSE

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A100   = new C_TB_1	' -- W8101MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
        cDataExists.A100.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A100.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지			
						
      
						
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제1호 서식", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
          
		    chkData1 =cDataExists.A100.W2                 '법인종류별구분 
		
		End If	
						
		
		Set cDataExists.A100 = Nothing
		
		



		chkData2 =lgcCompanyInfo.COMP_TYPE1							'법인구분 
		chkData3 =lgcCompanyInfo.HOME_TAX_MAIN_IND					'주종목번호 
		chkData4 =mid(replace(lgcCompanyInfo.OWN_RGST_NO,"-",""),4,2)'사업자번호(4,2)
					
		

		
		Set cDataExists = Nothing	' -- 메모리해제 
		

		
	   If (chkData1 = "60" Or chkData1 = "70" ) and chkData3 = "999999" then
	      CHK_COMPANY = False
	      Exit Function
	   End if   
	   
	   
	     If chkData2 = "2" and  chkData4 = "84" then
	      CHK_COMPANY = False
	      Exit Function
	   End if 

 CHK_COMPANY = TRUE
End Function

Function Getdata_TB_3_3_4_A142(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A113   = new C_TB_3_2	' -- W1101MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
        cDataExists.A113.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A113.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지			
		
		
	
		   Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   '일반법인 
		  	
		   cDataExists.A113.WHERE_SQL  =  lgStrSQL			
		
		If Not cDataExists.A113.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " 표준대차대조표", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
          
		    dblData =unicdbl(cDataExists.A113.CR_INV,0)          
		
		End If	
						
		
		Set cDataExists.A113 = Nothing
		
		
		
		Set cDataExists = Nothing	' -- 메모리해제 
		
		 Getdata_TB_3_3_4_A142 = dblData
	
End Function



Function Getdata_TB_3_3_A142(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A115   = new C_TB_3_3	' -- W1101MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
        cDataExists.A115.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A115.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지			
		
		
		   Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   '일반법인 
		  				

		   cDataExists.A115.WHERE_SQL  =  lgStrSQL				
		If Not cDataExists.A115.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " 표준손익계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
          
		    dblData =unicdbl(cDataExists.A115.W5,0)          
		
		End If	
						
		
		Set cDataExists.A115 = Nothing
		
		
		
		Set cDataExists = Nothing	' -- 메모리해제 
		
		 Getdata_TB_3_3_A142 = dblData
	
End Function

' -- 2006.03 개정추가 
Function Getdata_TB_47_A110(byVal pType )
 Dim chkData,iKey1, iKey2, iKey3,cDataExists , dblData

        
		Set cDataExists = new TYPE_DATA_EXIST_W1119MA1
		Set cDataExists.A110   = new C_TB_47A	' -- W9101MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
        cDataExists.A110.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A110.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지			
		
		Call SubMakeSQLStatements_W1119MA1(pType,iKey1, iKey2, iKey3) 	   
		  	
		   cDataExists.A110.WHERE_SQL  =  lgStrSQL			
		
		If Not cDataExists.A110.LoadData() Then
			'blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제47호 주요계정명세서(갑) ", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
          
		    dblData =unicdbl(cDataExists.A110.W125,0)          
		
		End If	
						
		
		Set cDataExists.A110 = Nothing
		
		Set cDataExists = Nothing	' -- 메모리해제 
		
		 Getdata_TB_47_A110 = dblData
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W1119MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A113_1" '-- 외부 참조 SQL
	       lgStrSQL =""
	       lgStrSQL = " and  par_gp_cd = '33'  and A.W4 = '56'  "
	  
	  
	  Case "A114_1" '-- 외부 참조 SQL
	       lgStrSQL =""
	       lgStrSQL = " and par_gp_cd = '33'  and  A.W4 = '78' "    
	       
	  Case "A115_1" '-- 외부 참조 SQL
	       lgStrSQL =""
	       lgStrSQL = " and A.W4 = '82'  "
	  
	  
	  Case "A116_1" '-- 외부 참조 SQL
	       lgStrSQL =""
	       lgStrSQL = " and  A.W4 = '73' "            
   
	  Case "A110"
		   lgStrSQL =""
	End Select
	PrintLog "SubMakeSQLStatements_W1119MA1 : " & lgStrSQL
End Sub
%>
