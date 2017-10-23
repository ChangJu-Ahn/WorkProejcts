<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 조특제2호2 기업의어음제도개선세액공제 
'*  3. Program ID           : W6109MA1
'*  4. Program Name         : W6109MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_JT2_2

Set lgcTB_JT2_2 = Nothing	' -- 초기화 

Class C_TB_JT2_2
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1
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

		lgStrSQL = ""
		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgoRs1,lgStrSQL, "", "") = False Then
			If Not CALLED_OUT Then	' -- 외부에서 부른 경우는 호출한쪽에서 데이타없음을 저정한다. 이유는 lgsPGM_ID, lgsPGM_NM가 호출한넘이기때문이다.
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		
		LoadData = True
	End Function

	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData = lgoRs1(pFieldNm)
		End If
	End Function
	
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)	
	End Sub	
	
	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf

	            If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보의 업종을 불러온다 
					lgStrSQL = lgStrSQL & " , B.IND_TYPE " & vbCrLf
	            End If
	            	            
				lgStrSQL = lgStrSQL & " FROM TB_JT2_2_200603	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 

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
Class TYPE_DATA_EXIST_W6109MA1
	Dim A165

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6109MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6109MA1"
	
	Set lgcTB_JT2_2 = New C_TB_JT2_2		' -- 해당서식 클래스 
	
	If Not lgcTB_JT2_2.LoadData	Then Exit Function		' -- 제1호 서식 로드 
	
	Set cDataExists = new  TYPE_DATA_EXIST_W6109MA1	
	
	'==========================================
	' -- 조특제2호2 기업의어음제도개선세액공제 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("IND_TYPE"), "법인정보_업종") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_JT2_2.GetData("IND_TYPE"), 50)

' -- 서식개정으로 삭제됨 : 200603 
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1"), "환어음.판매대금추심의뢰서 결제금액및 기업구매전용카드 사용금액및 외상매출채권담보대출제도 이용금액") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1"), 15, 0)
'	if  UNICDbl(lgcTB_JT2_2.GetData("W1"),  0) <> UNICDbl(lgcTB_JT2_2.GetData("W1_A"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_B"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_C"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W1_D"),0) then
'	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "환어음.판매대금추심의뢰서 결제금액및 기업구매전용카드 사용금액및 외상매출채권담보대출제도 이용금액","환어음결제금액 + 판매대금추심의뢰서결제금액 + 기업구매전용카드사용금액 +외상매출채권담보대출제도이용액"))
'	    blnError = True
'	end if
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_A"), "환어음결제금액") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_A"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_B"), "판매대금추심의뢰서결제금액") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_B"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_C"), "기업구매전용카드사용금액") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_C"), 15, 0)
'	
'	If Not ChkNotNull(lgcTB_JT2_2.GetData("W1_D"), "외상매출채권담보대출제도이용금액") Then blnError = True		
'	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W1_D"), 15, 0)

'	'공제금액 : (항목(8)환어음판매대금추심의뢰서결제금액,기업구매전용카드사용금액 및외상매출채권담보제도 이용금액 - 항목(9)약속어음결제금액) x 3 /1000 
'	If  ChkNotNull(lgcTB_JT2_2.GetData("W3"), "공제금액") Then 
'	    if  UNICDbl(lgcTB_JT2_2.GetData("W3"),  0)  <> Fix((UNICDbl(lgcTB_JT2_2.GetData("W1"), 0) - UNICDbl(lgcTB_JT2_2.GetData("W2"),0))* UNICDbl(lgcTB_JT2_2.GetData("W3_RATE_VALUE"),0))  then
'	         Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.W3, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제금액","  (항목(8)환어음판매대금추심의뢰서결제금액,기업구매전용카드사용금액 및외상매출채권담보제도 이용금액 - 항목(9)약속어음결제금액) x 3 /1000(소수점절사) "))
'	         blnError = True
'	    End if
'	Else	
'		blnError = True		
'	End if	
'--------------------------------------
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W11"), "약속어음결제금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W11"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_HAP_C"), "공제금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_HAP_C"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W13"), "산출세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W13"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W14"), "한도액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W14"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W15"), "공제세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W15"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_SUM"), "대상금액_합계") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_SUM"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_1"), "대상금액_지급기한1합계") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_HAP_2"), "대상금액_지급기한2합계") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_HAP_2"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_GA_C"), "공제대상금액_가") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_GA_C"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W12_NA_C"), "공제대상금액_나") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W12_NA_C"), 15, 0)

	sHTFBody = sHTFBody & UNIChar("", 44)	' -- 공란 


	' -- 점검 : 공제금액 = a * b 의 Sum
	if  UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W12_GA_C"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W12_NA_C"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W12_GA_C"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(12) 공제금액"," 공제금액( (a) X (b) )의 가 + 나"))
	     blnError = True
	end if
	
	' -- 점검 : (15)공제새액은 (12)와 (14)중 적은 금액 
	If UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) < UNICDbl(lgcTB_JT2_2.GetData("W14"), 0) Then
		if  UNICDbl(lgcTB_JT2_2.GetData("W12_HAP_C"), 0) <> UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15) 공제세액","(12)공제세액"))
		     blnError = True
		end if
	Else
		if  UNICDbl(lgcTB_JT2_2.GetData("W14"), 0) <> UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) then
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15) 공제세액","(14)한도액"))
		     blnError = True
		end if
	End If
	
	
	' -- 점검 : W15 금액과 일치여부 
	If  UNICDbl(lgcTB_JT2_2.GetData("W15"), 0) > 0 Then 
	
		Set cDataExists.A165  = new C_TB_JT1	' -- W6103MA1_HTF.asp 에 정의됨 
		
		Call SubMakeSQLStatements_W6109MA1("A165",iKey1, iKey2, iKey3)   
		
		cDataExists.A165.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A165.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
	
	   ' 공제세액(13) : 항목(10)공제금액과 항목(12)한도액중에 적은금액을 공제세액에 입력합니다.
       ' 세액공제신청서(A165)의 코드(75) 기업의어음제도개선을위한 세액공제 항목(11) 대상세액  과 일치 

        If Not cDataExists.A165.LoadData then
            blnError = True
			Call SaveHTFError(lgsPGM_ID, "조특 제 1호  세액공제신청서(A165)", TYPE_DATA_NOT_FOUND)	
        else

		  	if UNICDBL(lgcTB_JT2_2.GetData("W15"),  0) <> UNICDbl(cDataExists.A165.GetData("W5"), 0) then

		  	    Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15)공제세액 "," 세액공제신청서(A165)의 코드(75) 기업의어음제도개선을위한 세액공제 항목(6) 공제세액"))
				blnError = True
				
		  	end if
		  	
		  	
		End if   
		Set cDataExists.A165 = Nothing	   
	Else
	    blnError = True
	End if    		



	
	' -- 기업의 어음제도개선을 위한 공제새액계산서 - 대상금액 
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "1"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_SUM"), "환어음 결제금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_1"), "환어음 결제금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_A_2"), "환어음 결제금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_A_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 
	
	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_A_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_A_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_A_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_A_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "환어음 결제금액 합계","환어음 결제금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if
	
	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "2"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_SUM"), "판매대금추심의뢰서 결제금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_1"), "판매대금추심의뢰서 결제금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_B_2"), "판매대금추심의뢰서 결제금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_B_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 

	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_B_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_B_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_B_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_B_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "판매대금추심의뢰서 결제금액 합계","판매대금추심의뢰서 결제금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "3"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_SUM"), "기업구매전용카드 사용금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_1"), "기업구매전용카드 사용금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_C_2"), "기업구매전용카드 사용금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_C_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 

	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_C_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_C_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_C_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_C_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기업구매전용카드 사용금액 합계","기업구매전용카드 사용금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "4"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_SUM"), "외상매출채권담보대출제도 이용금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_1"), "외상매출채권담보대출제도 이용금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_D_2"), "외상매출채권담보대출제도 이용금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_D_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 

	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_D_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_D_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_D_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_E_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "외상매출채권담보대출제도 이용금액 합계","외상매출채권담보대출제도 이용금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "5"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_SUM"), "구매론제도 이용금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_1"), "구매론제도 이용금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_E_2"), "구매론제도 이용금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_E_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 

	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_E_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_E_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_E_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_E_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "구매론제도 이용금액 합계","구매론제도 이용금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if


	sHTFBody = sHTFBody & "84"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	sHTFBody = sHTFBody & "6"

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_SUM"), "네트워크론제도 이용금액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_SUM"), 15, 0)
	
	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_1"), "네트워크론제도 이용금액 30일 이내") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_1"), 15, 0)

	If Not ChkNotNull(lgcTB_JT2_2.GetData("W8_F_2"), "네트워크론제도 이용금액 31일 ~ 60일") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_JT2_2.GetData("W8_F_2"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 48)	' -- 공란 


	' -- 점검 : 대상금액 합계 = 30일 이내 + 31일 ~ 60일 
	if  UNICDbl(lgcTB_JT2_2.GetData("W8_F_SUM"), 0) <> ( UNICDbl(lgcTB_JT2_2.GetData("W8_F_1"), 0) + UNICDbl(lgcTB_JT2_2.GetData("W8_F_2"), 0) ) then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JT2_2.GetData("W8_F_SUM"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "네트워크론제도 이용금액 합계","네트워크론제도 이용금액의 지급기한 30일이내 + 31일~60일"))
	     blnError = True
	end if

	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_JT2_2 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	 Case "A165" '-- 외부 참조 SQL	
	      ' 세액공제신청서(A165)의 코드(75) 기업의어음제도개선을위한 세액공제 항목(11) 대상세액과 일치 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " and A.W3 = '75'" & vbCrLf
	
	End Select
	PrintLog "SubMakeSQLStatements_W6109MA1 : " & lgStrSQL
End Sub

%>
