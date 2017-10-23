
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제1호 법인세과세표준 신고서 
'*  3. Program ID           : W8107MA1
'*  4. Program Name         : W8107MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_1

Set lgcTB_1 = Nothing	' -- 초기화 

Class C_TB_1
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W1_RATE
	Dim W1_RATE_View
	Dim W2
	Dim W2_A
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim W8
	Dim W8_A
	Dim W9
	Dim W10
	Dim W11
	Dim W12_A
	Dim W12_B
	Dim W13
	Dim W14
	Dim W15
	Dim W16
	Dim W17_1
	Dim W17_2
	Dim W17_Sum
	Dim W18_1
	Dim W18_2
	Dim W18_Sum
	Dim W19_1
	Dim W19_2
	Dim W19_Sum
	Dim W20_1
	Dim W20_2
	Dim W20_Sum
	Dim W21_1
	Dim W21_2
	Dim W21_Sum
	Dim W22_1
	Dim W23_1
	Dim W_TYPE
		
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, oRs1
			 
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

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
				  
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W1				= oRs1("W1")
		W1_RATE			= oRs1("W1_RATE")
		W1_RATE_VIEW	= oRs1("W1_RATE_VIEW")
		W2				= oRs1("W2")
		W2_A			= oRs1("W2_A")
		W3				= oRs1("W3")
		W4				= oRs1("W4")
		W5				= oRs1("W5")
		W6				= oRs1("W6")
		W7				= oRs1("W7")
		W8				= oRs1("W8")
		W8_A			= oRs1("W8_A")
		W9				= oRs1("W9")
		W10				= oRs1("W10")
		W11				= oRs1("W11")
		W12_A			= oRs1("W12_A")
		W12_B			= oRs1("W12_B")
		W13				= oRs1("W13")
		W14				= oRs1("W14")
		W15				= oRs1("W15")
		W16				= oRs1("W16")
		W17_1			= oRs1("W17_1")
		W17_2			= oRs1("W17_2")
		W17_SUM			= oRs1("W17_SUM")
		W18_1			= oRs1("W18_1")
		W18_2			= oRs1("W18_2")
		W18_SUM			= oRs1("W18_SUM")
		W19_1			= oRs1("W19_1")
		W19_2			= oRs1("W19_2")
		W19_SUM			= oRs1("W19_SUM")
		W20_1			= oRs1("W20_1")
		W20_2			= oRs1("W20_2")
		W20_SUM			= oRs1("W20_SUM")
		W21_1			= oRs1("W21_1")
		W21_2			= oRs1("W21_2")
		W21_SUM			= oRs1("W21_SUM")
		W22_1			= oRs1("W22_1")
		W23_1			= oRs1("W23_1")
		W_TYPE			= oRs1("W_TYPE")
		
		Call SubCloseRs(oRs1)	
		
		LoadData = True
	End Function

	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub	
	
	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_1	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W8107MA1
	Dim A131
	Dim A132
	Dim A130
	Dim A170
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8107MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8107MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8107MA1"
	
	Set lgcTB_1 = New C_TB_1		' -- 해당서식 클래스 
	
	If Not lgcTB_1.LoadData	Then Exit Function		' -- 제1호 서식 로드 
		
	'==========================================
	' -- 제1호 법인세과세표준신고서 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkBoundary("1,2,3", lgcTB_1.W1, "법인구분: " & lgcTB_1.W1 & " " ) Then blnError = True
    sHTFBody = sHTFBody & UNIChar(lgcTB_1.W1, 1)
	If lgcTB_1.W1 = "3" Then
		If UNICDbl(lgcTB_1.W1_RATE, 0) = 0 Then	' 코드 3 일때 0보다 커야 한 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W1, UNIGetMesg(TYPE_CHK_ZERO_OVER, "외투 비율",""))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W1_RATE, 6, 3)
	Else
		sHTFBody = sHTFBody & UNINumeric(0, 6, 3)
	End If
	
	' 사업자번호(4:2)가 '84'인데 법인구분이 '2'가 아니면 오류. ('3'인데 '84'도 있음)
	If (lgcTB_1.W1 <> "2" OR lgcTB_1.W1 <> "3" ) And GetRgstNo42(lgcCompanyInfo.OWN_RGST_NO) = "84"    Then	
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W1, UNIGetMesg("사업자번호(4:2)가 '84'인데 법인구분이 '2'가 아니면 오류. ('3'인데 '84'도 있음)", "",""))
	End If
	
	If Not ChkBoundary("11,12,21,22,30,40,50,60,70", lgcTB_1.W2, "종류별구분: " & lgcTB_1.W2 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W2, 2)

	If Not ChkBoundary("1,2", lgcTB_1.W3, "조정구분: " & lgcTB_1.W3 & " " ) Then blnError = True
	If lgcCompanyInfo.EX_RECON_FLG	= "Y" And lgcTB_1.W3 = "2" Then		' 법인정보에 외부조정인데, 자기조정이라고 체크하면 에러 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "조정구분: 자기", UNIGetMesg("법인기초정보관리의 외부조정여부가 '예' 입니다", "",""))
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W3, 1)
	

	
	If Not ChkBoundary("1,2", lgcTB_1.W4, "외부감사여부: " & lgcTB_1.W4 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W4, 1)
	
	If Not ChkNotNull(lgcTB_1.W5, "결산확정일") Then blnError = True
	If DateDiff("m", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W5) < 0 Then 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W5, UNIGetMesg("당기종료일자보다 작습니다.", "",""))
	End If
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W5)
	
	If Not ChkNotNull(lgcTB_1.W6, "신고일") Then blnError = True
	If DateDiff("d", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W6) < 0 Or _
	   DateDiff("m", lgcCompanyInfo.FISC_END_DT, lgcTB_1.W6) > 3  Then 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W6, UNIGetMesg("당기종료일자보다 작거나, 3개월을 초과하였습니다", "",""))
	ElseIf DateDiff("d", Date(), lgcTB_1.W6 ) < 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W6, UNIGetMesg("신고일이 현재보다 이전입니다", "",""))
	End If
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W6)


	sHTFBody = sHTFBody & "10"	' -- 신고구분 

	' ------------- 데이타 존재 체크를 위해 -------------------
	Set cDataExists	= new TYPE_DATA_EXIST_W8107MA1

	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 
		
	' --- 데이타 조회 SQL
	Call SubMakeSQLStatements_W8107MA1("O",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,oRs2,lgStrSQL, "", "") = False Then
		blnError = True
	    'Call SaveHTFError(lgsPGM_ID, "조정후수입금액명세서(A111)", TYPE_DATA_NOT_FOUND)
	    'Call SaveHTFError(lgsPGM_ID, "법인과세표준및세액조정계산서(A101)", TYPE_DATA_NOT_FOUND)
	
	Else
		If oRs2("W_TYPE") = "0" Then
			cDataExists.A130 = "" & oRs2("W_21")	' 서식 존재 체크 0: 없음, 1: 존재 
			cDataExists.A131 = "" & oRs2("W_19")
			cDataExists.A132 = "" & oRs2("W_20")
			cDataExists.A170 = "" & oRs2("W_22")
		Else
			cDataExists.A130 = "0" & oRs2("W_21")	
			cDataExists.A131 = "0" & oRs2("W_19")
			cDataExists.A132 = "0" & oRs2("W_20")
			cDataExists.A170 = "0" & oRs2("W_22")
		End If
		oRs2.MoveNext		' W_TYPE 증가를 위해 다음 레코드로 넘김 
	End If

	'PrintLog "----------.. : " & sHTFBody
	If Not ChkBoundary("1,2", lgcTB_1.W9, "주식변동여부: " & lgcTB_1.W9 & " " ) Then blnError = True
	If lgcTB_1.W9 = "1" Then
		If cDataExists.A131 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W9, UNIGetMesg("주식등변동상황명세서(A131)자료가 없습니다", "",""))
		End If
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W9, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W10, "장부전산화여부: " & lgcTB_1.W10 & " " ) Then blnError = True
	If lgcTB_1.W10 = "1" Then
		If cDataExists.A130 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W10, UNIGetMesg("전산조직운용명세서(A130)자료가 없습니다", "",""))
		End If	
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W10, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W11, "사업년도의제여: " & lgcTB_1.W11 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W11, 1)

	If Not ChkDate(lgcTB_1.W12_A, "신고기간연장승인 - 신청") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W12_A)

	If Not ChkDate(lgcTB_1.W12_B, "신고기간연장승인 - 연장기한") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcTB_1.W12_B)

	If Not ChkBoundary("1,2", lgcTB_1.W13, "결손금소급공제 법인세환급신청여: " & lgcTB_1.W13 & " " ) Then blnError = True
	If lgcTB_1.W13 = "1" Then
		If cDataExists.A170 = "0" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W13, UNIGetMesg("소급공제법인세액환급신청서(A170)자료가 없습니다", "",""))
		End If	
	End If 
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W13, 1)
 
	If Not ChkBoundary("1,2", lgcTB_1.W14, "감가상각방법신고서세출여부: " & lgcTB_1.W14 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W14, 1)

	If Not ChkBoundary("1,2", lgcTB_1.W15, "재고자산등 평가방법신고서 제출여부: " & lgcTB_1.W15 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W15, 1)
		
	If UNICDbl(lgcTB_1.W16, 0) < 0 Then	' -- 수입금액 음수 오류체크 
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_1.W16, UNIGetMesg(TYPE_CHK_ZERO_OVER, "수입금액",""))
	End If
	
	If oRs2("W_TYPE") = "1" Then
		'-- 조정후수입금액명세서(A111)의 코드(99)수입금액 합계의  항목(4)계와 일치하지 않으면 오류 
		'--  (조정후수입금액명세서(A111)상의 수입금액이 “0”보다 큰 경우) 

		sTmp = UNICDbl(oRs2("W_19"), 0)

		If UNICDbl(lgcTB_1.W16, 0) <> sTmp Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W16, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "수입금액","조정후수입금액명세서(A111)의 코드(99)수입금액 합계의 항목(4)계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W16, 15, 0)
			
		oRs2.MoveNext	' -- 다음레코드 
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "조정후수입금액명세서(A111)", TYPE_DATA_NOT_FOUND)
	End If
	
	

	If oRs2("W_TYPE") = "2" Then
		' 20번 과세표준_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준과 일치 
		' 21번 과세표준_토지등 양도소득에 대한 법인 : 법인세과세표준및세액조정계산서(A101)의 코드(34)과세표준과 일치 
		' 22번 과세표준_계: 과세표준_법인세(31) + 과세표준_토지등 양도소득에 대한 법인세(31) 
		If UNICDbl(lgcTB_1.W17_1, 0) <> UNICDbl(oRs2("W_20"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_1 & " <> " & oRs2("W_20") , UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준_법인세","법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W17_2, 0) <> UNICDbl(oRs2("W_21"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준_토지등 양도소득에 대한 법인세","법인세과세표준및세액조정계산서(A101)의 코드(34)과세표준"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_2, 15, 0)
			
		
		
		If UNICDbl(lgcTB_1.W17_SUM, 0) <> UNICDbl(lgcTB_1.W17_1, 0)  +  UNICDbl(lgcTB_1.W17_2, 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W17_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준_계","과세표준_법인세 + 과세표준_토지등 양도소득에 대한 법인세"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W17_SUM, 15, 0)
			
		'23. 산출세액_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(16)산출세액_합계와 일치 
		'24. 산출세액_토지등 양도소득에 대한 법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(36)산출세액과 일치 
		'25. 산출세액_계	: 산출세액_법인세(32)  +  산출세액_토지 등 양도소득에 대한 법인세(32)
		If UNICDbl(lgcTB_1.W18_1, 0) <> UNICDbl(oRs2("W_23"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "산출세액_법인세","법인세과세표준및세액조정계산서(A101)의 코드(16)산출세액_합계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W18_2, 0) <> UNICDbl(oRs2("W_24"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "산출세액_토지등 양도소득에 대한 법인세","법인세과세표준및세액조정계산서(A101)의 코드(36)산출세액"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W18_SUM, 0) <> UNICDbl(lgcTB_1.W18_1, 0)  + UNICDbl(lgcTB_1.W18_2, 0)  Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W18_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "산출세액_계", "산출세액_법인세 + 산출세액_토지등 양도소득에 대한 법인세"))
		End If	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W18_SUM, 15, 0)
					

		'26. 총부담세액_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(21)납부할세액_가감계 + 코드(29)납부할세액_감면분추가납부 
		'27. 총부담세액_토지등 양도소득에 대한 법인세: 법인세과세표준및세액조정계산서(A101)의 코드(41)양도소득법인세_가감계와 일치 
		'28. 총부담세액_계 : 총부담세액_법인세(33) + 총부담세액_토지등 양도소득에 대한 법인세(33)			
		If UNICDbl(lgcTB_1.W19_1, 0) <> UNICDbl(oRs2("W_26"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "총부담세액_법인세","법인세과세표준및세액조정계산서(A101)의 코드(21)납부할세액_가감계 + 코드(29)납부할세액_감면분추가납부"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W19_2, 0) <> UNICDbl(oRs2("W_27"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "총부담세액_토지등 양도소득에 대한 법인세","법인세과세표준및세액조정계산서(A101)의 코드(41)양도소득법인세_가감계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W19_SUM, 0) <> UNICDbl(lgcTB_1.W19_1, 0) + UNICDbl(lgcTB_1.W19_2, 0)   Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W19_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "총부담세액_계","총부담세액_법인세 + 총부담세액_토지등 양도소득에 대한 법인세"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W19_SUM, 15, 0)
			
	
		'29. 기납부_법인세 : 법인과세표준및세액조정계산서(A101)의 코드(28)법인세 기납부세액_합계와 일치 
		'30. 기납부세액_토지등 양도소득에 대한 법인세: 법인과세표준및세액조정계산서(A101)의 코드(44)양도소득 기납부세액_계와 일치 
		'31. 기납부세액_계 : 기납부세액_법인세(34) + 기납부세_액토지등 양도소득에 대한 법인세(34)
		If UNICDbl(lgcTB_1.W20_1, 0) <> UNICDbl(oRs2("W_29"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기납부_법인세","법인과세표준및세액조정계산서(A101)의 코드(28)법인세 기납부세액_합계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W20_2, 0) <> UNICDbl(oRs2("W_30"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기납부세액_토지등 양도소득에 대한 법인세","법인과세표준및세액조정계산서(A101)의 코드(44)양도소득 기납부세액_계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W20_SUM, 0) <> UNICDbl(lgcTB_1.W20_1, 0) + UNICDbl(lgcTB_1.W20_2, 0)   Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W20_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기납부세액_계","기납부_법인세 + 기납부세액_토지등 양도소득에 대한 법인세"))
		End If		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W20_SUM, 15, 0)
				
			
		'32. 차감납부할세액_법인세	: 법인과세표준및세액조정계산서(A101)의 코드(30)법인세 차감납부할세액과 일치 
		'33. 차감납부할세액_토지등양도소득에 대한 법인세: 법인과세표준및세액조정계산서(A101)의 코드(45)양도소득 차감납부할세액과 일치 
		'34. 차감납부할세액_계 : (35)차감납부할세액_법인세 + (35)차감납부할세액_토지등 양도소득에 대한 법인세 
		If UNICDbl(lgcTB_1.W21_1, 0) <> UNICDbl(oRs2("W_32"), 0) Then
		
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부할세액_법인세","법인과세표준및세액조정계산서(A101)의 코드(28)법인세 기납부세액_합계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W21_2, 0) <> UNICDbl(oRs2("W_33"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_2, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부할세액_토지등양도소득에 대한 법인세","법인과세표준및세액조정계산서(A101)의 코드(45)양도소득 차감납부할세액"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_2, 15, 0)
			
		If UNICDbl(lgcTB_1.W21_SUM, 0) <> UNICDbl(lgcTB_1.W21_1, 0) + UNICDbl(lgcTB_1.W21_2, 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W21_SUM, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부할세액_계","차감납부할세액_법인세 + 차감납부할세액_토지등양도소득에 대한 법인세"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W21_SUM, 15, 0)
								

		'35. 분납할세액	: 법인세과세표준및세액조정계산서(A101)의 코드(50)분납할세액_계와 일치 
		'36. 차감납부세액 : 법인세과세표준및세액조정계산서(A101)의 코드(53)차감납부세액_계와 일치 
		If UNICDbl(lgcTB_1.W22_1, 0) <> UNICDbl(oRs2("W_35"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W22_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "분납할세액","법인세과세표준및세액조정계산서(A101)의 코드(50)분납할세액_계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W22_1, 15, 0)
			
		If UNICDbl(lgcTB_1.W23_1, 0) <> UNICDbl(oRs2("W_36"), 0) Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_1.W23_1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감납부세액","법인세과세표준및세액조정계산서(A101)의 코드(53)차감납부세액_계"))
		End If
		sHTFBody = sHTFBody & UNINumeric(lgcTB_1.W23_1, 15, 0)
			
	Else
		blnError = True
		Call SaveHTFError(lgsPGM_ID, "법인과세표준및세액조정계산서(A101)", TYPE_DATA_NOT_FOUND)
	End If
	
	' -- 조정구분에서 외부조정의 이면 체크로직을 가동후 데이타 추출/ 자기조정이면 데이타만 추출 
	If lgcTB_1.W3 = "1" Then
		If Not ChkNotNull(lgcCompanyInfo.RECON_BAN_NO, "외부조정자_조정반번호") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.RECON_MGT_NO, "외부조정자_조정자관리번호") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.AGENT_NM, "외부조정자_성명") Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.AGENT_RGST_NO, "외부조정자_사업자등록번호") Then blnError = True
	End If
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_BAN_NO), 5)
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_MGT_NO), 6)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_NM, 30)
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.AGENT_RGST_NO), 10)
	
	If Trim(lgcCompanyInfo.BANK_CD) <> "" Then ' 은행코드가 존재시 코드검증및 계좌번호입력체크 
		If Not ChkBoundary("02,03,05,06,07,10,11,12,13,14,15,20,21,23,26,27,31,32,34,35,37,39,71,72,73,74,75,81", lgcCompanyInfo.BANK_CD, "국세환급계좌_은행코드: " & lgcCompanyInfo.BANK_CD & " " ) Then blnError = True
		If Not ChkNotNull(lgcCompanyInfo.BANK_ACCT_NO, "국세환급계좌_계좌번호") Then blnError = True
	End If
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_CD, 2)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_DPST, 20)
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.BANK_ACCT_NO, 20)
	
	If Not ChkBoundary("Y,N", lgcCompanyInfo.EX_54_FLG, "주식변동자료매체로제출여부: " & lgcCompanyInfo.EX_54_FLG & " " ) Then blnError = True
	If lgcCompanyInfo.EX_54_FLG = "Y" Then	' -- Y 시 A131, A132 데이타 존재시 오류 
		If cDataExists.A131 = "1" Or cDataExists.A132 = "1" Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcCompanyInfo.EX_54_FLG, UNIGetMesg("법인기초정보_주식변동자료매체로제출여부가 '예'일 경우 주식등변동상황명세서(A131)자료가 존재하면 오류입니다", "",""))
		End If
	End If
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.EX_54_FLG, 1)

	If Not ChkNotNull(lgcTB_1.W_TYPE, "유형별 구분") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_1.W_TYPE, 3)
	
	sHTFBody = sHTFBody & UNIChar("", 26)	' -- 공란 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.

	If Not blnError Then

		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	'Set lgcTB_1 = Nothing	' -- 메모리해제  <-- W8101MA1_HTF에서 사용함 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8107MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
			
			lgStrSQL = ""
			' -- 검증을 위해 데이타 존재 체크 
			lgStrSQL = lgStrSQL & " SELECT '0' W_TYPE " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_54H WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_19 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_54_BPH WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_20 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_JS1 WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_21 " & vbCrLf
			lgStrSQL = lgStrSQL & "	,	ISNULL(( SELECT TOP 1 1 FROM TB_68 WITH (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		),0)  W_22 " & vbCrLf
			lgStrSQL = lgStrSQL & "		, 0 W_23, 0 W_24, 0 W_25, 0 W_26, 0 W_27, 0 W_28, 0 W_29" & vbCrLf		
			lgStrSQL = lgStrSQL & "		, 0 W_30, 0 W_31, 0 W_32, 0 W_33, 0 W_34, 0 W_35, 0 W_36" & vbCrLf	
			
			' -- 17호 서식 값			
			lgStrSQL = lgStrSQL & " UNION "		 					 & vbCrLf
			lgStrSQL = lgStrSQL & " SELECT '1' W_TYPE, A.W4 W_19" & vbCrLf
			lgStrSQL = lgStrSQL & "		, 0 W_20, 0 W_21, 0 W_22, 0 W_23, 0 W_24, 0 W_25, 0 W_26, 0 W_27, 0 W_28, 0 W_29" & vbCrLf		
			lgStrSQL = lgStrSQL & "		, 0 W_30, 0 W_31, 0 W_32, 0 W_33, 0 W_34, 0 W_35, 0 W_36" & vbCrLf		
			lgStrSQL = lgStrSQL & " FROM TB_17_D1	A " & vbCrLf	
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.CODE_NO = '99'"		 	 & vbCrLf

			' -- 데이타 검증을 위해 호출 
			' 20번 과세표준_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준과 일치 
			' 21번 과세표준_토지등 양도소득에 대한 법인 : 법인세과세표준및세액조정계산서(A101)의 코드(34)과세표준과 일치 
			' 22번 과세표준_계: 과세표준_법인세(31) + 과세표준_토지등 양도소득에 대한 법인세(31) 
			lgStrSQL = lgStrSQL & " UNION "		 					 & vbCrLf
			'lgStrSQL = lgStrSQL & " SELECT	'2' W_TYPE, 0 W_19, A.W10 W_20, A.W34 W_21, A.W31 + A.W34 W_22" & vbCrLf		
			lgStrSQL = lgStrSQL & " SELECT	'2' W_TYPE, 0 W_19, A.W56 W_20, A.W34 W_21, A.W31 + A.W34 W_22" & vbCrLf		' W10 => W56 : 200603 개정 
			
			'23. 산출세액_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(16)산출세액_합계와 일치 
			'24. 산출세액_토지등 양도소득에 대한 법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(36)산출세액과 일치 
			'25. 산출세액_계	: 산출세액_법인세(32)  +  산출세액_토지 등 양도소득에 대한 법인세(32)
			lgStrSQL = lgStrSQL & "		,	A.W16 W_23, A.W36 W_24, A.W16 + A.W36 W_25" & vbCrLf		

			'26. 총부담세액_법인세 : 법인세과세표준및세액조정계산서(A101)의 코드(21)납부할세액_가감계 + 코드(29)납부할세액_감면분추가납부 
			'27. 총부담세액_토지등 양도소득에 대한 법인세: 법인세과세표준및세액조정계산서(A101)의 코드(41)양도소득법인세_가감계와 일치 
			'28. 총부담세액_계 : 총부담세액_법인세(33) + 총부담세액_토지등 양도소득에 대한 법인세(33)
			lgStrSQL = lgStrSQL & "		,  A.W21 + A.W29 W_26, A.W41 W_27, A.W21 + A.W29 + A.W41 W_28" & vbCrLf		

			'29. 기납부_법인세 : 법인과세표준및세액조정계산서(A101)의 코드(28)법인세 기납부세액_합계와 일치 
			'30. 기납부세액_토지등 양도소득에 대한 법인세: 법인과세표준및세액조정계산서(A101)의 코드(44)양도소득 기납부세액_계와 일치 
			'31. 기납부세액_계 : 기납부세액_법인세(34) + 기납부세_액토지등 양도소득에 대한 법인세(34)
			lgStrSQL = lgStrSQL & "		,  A.W28 W_29, A.W44 W_30, A.W29 + A.W44 W_31" & vbCrLf		
			
			'32. 차감납부할세액_법인세	: 법인과세표준및세액조정계산서(A101)의 코드(30)법인세 차감납부할세액과 일치 
			'33. 차감납부할세액_토지등양도소득에 대한 법인세: 법인과세표준및세액조정계산서(A101)의 코드(45)양도소득 차감납부할세액과 일치 
			'34. 차감납부할세액_계 : (35)차감납부할세액_법인세 + (35)차감납부할세액_토지등 양도소득에 대한 법인세 
			lgStrSQL = lgStrSQL & "		,  A.W30 W_32, A.W45 W_33, A.W32 + A.W45 W_34" & vbCrLf	
			
			'35. 분납할세액	: 법인세과세표준및세액조정계산서(A101)의 코드(50)분납할세액_계와 일치 
			'36. 차감납부세액 : 법인세과세표준및세액조정계산서(A101)의 코드(53)차감납부세액_계와 일치 
			lgStrSQL = lgStrSQL & "		,  A.W50 W_35, A.W53 W_36 " & vbCrLf	
			lgStrSQL = lgStrSQL & " FROM TB_3	A " & vbCrLf	
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			
	End Select
	PrintLog "SubMakeSQLStatements_W8107MA1 : " & lgStrSQL
End Sub

%>
