<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제68호 소급공제법인세액환급신청서 
'*  3. Program ID           : W8105MA1
'*  4. Program Name         : W8105MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_68

Set lgcTB_68 = Nothing	' -- 초기화 

Class C_TB_68
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
	            
	            If WHERE_SQL = "" Then	' 외부호출이 아니면 법인정보의 은행계좌 불러온다 
					lgStrSQL = lgStrSQL & " , B.BANK_CD, B.BANK_BRANCH, B.BANK_DPST, B.BANK_ACCT_NO " & vbCrLf
	            End If
	            
				lgStrSQL = lgStrSQL & " FROM TB_68	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				
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
Class TYPE_DATA_EXIST_W8105MA1
	Dim A101

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8105MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
  '  On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8105MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8105MA1"
	
	Set lgcTB_68 = New C_TB_68		' -- 해당서식 클래스 
	
	If Not lgcTB_68.LoadData	Then Exit Function		' -- 제1호 서식 로드 
	Set cDataExists = new  TYPE_DATA_EXIST_W8105MA1		
	'==========================================
	' -- 제68호 소급공제법인세액환급신청서 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_68.GetData("W1_S"), "결손사업연도_시작일") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W1_S"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W1_E"), "결손사업연도_종료") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W1_E"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W2_S"), "직전사업연도_시작일") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W2_S"))
	
	If Not ChkNotNull(lgcTB_68.GetData("W2_E"), "직전사업연도_종료") Then blnError = True		
	sHTFBody = sHTFBody & UNI8Date(lgcTB_68.GetData("W2_E"))
	
	If  ChkNotNull(lgcTB_68.GetData("W6"), "결손금액") Then 
	    
		Set cDataExists.A101  = new C_TB_3	' -- W8101MA1_HTF.asp 에 정의됨 
		Call SubMakeSQLStatements_W8105MA1("0",iKey1, iKey2, iKey3)   
		'cDataExists.A101.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		'cDataExists.A101.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
		
		 If Not cDataExists.A101.LoadData()  then
			    blnError = True
				Call SaveHTFError(lgsPGM_ID, "제 3호 법인세과세표준및세액조정계산서(A101)", TYPE_DATA_NOT_FOUND)	
         Else

				 if    UNICDbl(cDataExists.A101.W06, 0) >= 0  and UNICDbl(lgcTB_68.GetData("W6"),0)  <> 0 then
				     Call SaveHTFError(lgsPGM_ID,lgcTB_68.GetData("W6"), UNIGetMesg("신청대상이 아닙니다","", ""))
				     blnError = True	
				 Else    
						'결손금액 : 법인세과세표준및세액조정계산서(A101)의 코드(06)각사업년도소득금액 X (-1)과 같아야 함 
						if   UNICDbl(lgcTB_68.GetData("W6"),0) <> UNICDbl(cDataExists.A101.W06, 0) * -1 then
						     Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "결손금액", "제 3호 법인세과세표준및세액조정계산서(A101)의 코드(06)각사업년도소득금액 X (-1)"))
						    blnError = True		
						end if
        
        
				end if
		end if		
			Set cDataExists.A101 = Nothing	
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W6"), 15, 0)
	
	'소급공제받을 결손금액란： (8)란의 과세표준금액과 (6)란의 결손금액보다 작거나 같아야 함,
	If  ChkNotNull(lgcTB_68.GetData("W7"), "소급공제받을결손금액") Then
	    if UNICDbl(lgcTB_68.GetData("W7"),0) > UNICDbl(lgcTB_68.GetData("W8") ,0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W8"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "소급공제받을결손금액", "과세표준"))
	        blnError = True		
	    end if
	    
	    if UNICDbl(lgcTB_68.GetData("W7"),0) > UNICDbl(lgcTB_68.GetData("W6") ,0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W6"), UNIGetMesg(TYPE_CHK_LOW_EQUAL_AMT, "소급공제받을결손금액", "결손금액"))
	        blnError = True		
	    end if
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W7"), 15, 0)
	
	

	
	If Not ChkNotNull(lgcTB_68.GetData("W8"), "과세표준") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W8"), 15, 0)
	
	If Not ChkNotNull(Replace(lgcTB_68.GetData("W9"),"%",""), "세율") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(Replace(lgcTB_68.GetData("W9"),"%",""), 15, 0)
	
	If Not ChkNotNull(lgcTB_68.GetData("W10"), "산출세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W10"), 15, 0)
	
	If Not ChkNotNull(lgcTB_68.GetData("W11"), "공제감면세액") Then blnError = True		
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W11"), 15, 0)
	
	If ChkNotNull(lgcTB_68.GetData("W12"), "차감세액") Then
	   '차감세액 = 산출세액 - 공제감면세액 
	   if   UNICDbl(lgcTB_68.GetData("W12"),0) <>  UNICDbl(lgcTB_68.GetData("W10"),0) - UNICDbl(lgcTB_68.GetData("W11"),0) then
	        Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W12"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "차감세액", "산출세액 - 공제감면세액"))
	        blnError = True		
	   End if     
	Else 
		  blnError = True		
	End if   
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W12"), 15, 0)

	
	If  ChkNotNull(lgcTB_68.GetData("W13"), "직전사업연도법인세액") Then
	    If UNICDbl(lgcTB_68.GetData("W13"),0) <> UNICDbl(lgcTB_68.GetData("W10"),0) then
	       '직전사업연도법인세액: 항목 (10)산출세액과 같다 
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W13"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "직전사업연도법인세액", "산출세액"))
	       blnError = True		
	    End if
	Else
	      blnError = True		
	End if    
	    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W13"), 15, 0)
	'차감할세액 : 차감할세액 ≥ (산출세액 - 차감세액)
	If  ChkNotNull(lgcTB_68.GetData("W14"), "차감할세액") Then
	    if UNICDbl(lgcTB_68.GetData("W14"),0) <  UNICDbl(lgcTB_68.GetData("W10"),0) - UNICDbl(lgcTB_68.GetData("W12"),0) then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W14"), UNIGetMesg(TYPE_CHK_OVER_EQUAL, "차감할세액", "(산출세액 - 차감세액)"))
	   	        blnError = True	
	    End if
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W14"), 15, 0)
	
	
	'환급신청세액 : 직전사업연도법인세액 - 차감할세액 
	If  ChkNotNull(lgcTB_68.GetData("W15"), "환급신청세액") Then 
	    if UNICDbl(lgcTB_68.GetData("W15"),0) <> UNICDbl(lgcTB_68.GetData("W13"),0) - UNICDbl(lgcTB_68.GetData("W14"),0) Then
	       Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("W15"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "환급신청세액", "직전사업연도법인세액 - 차감할세액"))
	        blnError = True		
	    End if
	    
	Else
	    blnError = True		
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_68.GetData("W15"), 15, 0)
	
	if (Trim(lgcTB_68.GetData("BANK_CD")) = "" and Trim(lgcTB_68.GetData("BANK_ACCT_NO")) <> "")  Or (Trim(lgcTB_68.GetData("BANK_CD")) <> "" and Trim(lgcTB_68.GetData("BANK_ACCT_NO")) = "") Then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_68.GetData("BANK_CD") & Trim(lgcTB_68.GetData("BANK_ACCT_NO")), UNIGetMesg("은행코드 또는 계좌번호가 둘 중 하나만 입력되어 있습니다", "", ""))
	        blnError = True		
	End if
	
	' Null 허용 
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_CD"), "예입처(은행)코드") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_CD"), 2)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_BRANCH"), "예입처(본)지점") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_BRANCH"), 20)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_DPST"), "예금종류") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_DPST"), 20)
	
	'If Not ChkNotNull(lgcTB_68.GetData("BANK_ACCT_NO"), "계좌번호") Then blnError = True		
	sHTFBody = sHTFBody & UNIChar(lgcTB_68.GetData("BANK_ACCT_NO"), 20)
	

	sHTFBody = sHTFBody & UNIChar("", 60)	' -- 공란 
	
	' ----------- 
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_68 = Nothing	' -- 메모리해제  
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8105MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
	
	End Select
	PrintLog "SubMakeSQLStatements_W8105MA1 : " & lgStrSQL
End Sub

%>
