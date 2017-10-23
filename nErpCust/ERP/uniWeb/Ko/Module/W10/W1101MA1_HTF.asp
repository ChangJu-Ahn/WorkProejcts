
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제3호의2 표준대차대조표 
'*  3. Program ID           : W1101MA1
'*  4. Program Name         : W1101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_3_2

Set lgcTB_3_2 = Nothing ' -- 초기화 

Class C_TB_3_2
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	Dim DR_INV
	Dim CR_INV
	
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
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
			End If
		    Exit Function
		End If

		' 멀티행이지만 첫행을 리턴 
		Call GetData
		
		LoadData = True
	End Function

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.Find pWhereSQL
	End Function

	Function Filter(Byval pWhereSQL)
		lgoRs1.Filter = pWhereSQL
        Call GetData
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData
	End Function	
	
	Function Clone(Byref pRs)
		Set pRs = lgoRs1.clone
	End Function
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
			DR_INV		= lgoRs1("DR_INV")
			CR_INV		= lgoRs1("CR_INV")
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
				lgStrSQL = lgStrSQL & "SELECT A.W1, A.W2, ( CASE WHEN  LEFT(A.W2, 1) = '1' THEN B.W5 ELSE 0 END ) AS DR_INV, A.W5 " & vbCrLf
				lgStrSQL = lgStrSQL & ",  A.W3, A.W4, A.W6 " & vbCrLf
				lgStrSQL = lgStrSQL & ", (CASE WHEN LEFT(A.W2, 1) <> '1' THEN B.W5 ELSE 0 END ) AS CR_INV " & vbCrLf
				lgStrSQL = lgStrSQL & "FROM TB_3_2_2 A (NOLOCK)   " & vbCrLf
				lgStrSQL = lgStrSQL & "	INNER JOIN TB_3_2_1 B (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE AND A.W1=B.W1 AND A.W2=B.W2 " & vbCrLf 
				lgStrSQL = lgStrSQL & "	INNER JOIN TB_COMPANY_HISTORY C (NOLOCK) ON A.CO_CD=C.CO_CD AND A.FISC_YEAR=C.FISC_YEAR AND A.REP_TYPE=C.REP_TYPE AND A.W1=C.COMP_TYPE2  " & vbCrLf 
				lgStrSQL = lgStrSQL & "	LEFT OUTER JOIN dbo.ufn_TB_ACCT_GP('200503') D ON A.W1=D.COMP_TYPE2 AND A.W2=D.GP_CD  " & vbCrLf
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
				
				lgStrSQL = lgStrSQL & "  ORDER BY  LEFT(A.W2, 1), D.gp_seq " & vbCrLf
  				
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W1101MA1
	Dim A144

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W1101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, blnFirst, blnFile,dblAmt1 ,dblAmt72,dblAmt66,dblAmt43,dblAmt84
    
    Const TYPE_1 = 0	' 표준대차대조표(자산)
    Const TYPE_2 = 1	' 표준대차대조표(부채자본)
    Const TYPE_3 = 2	' 합계표준대차대조표(자산)
    Const TYPE_4 = 3	' 합계표준대차대조표(부채자본)
    Dim arrHTFBody(3)

    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False		' -- 에러 발생유무 
    blnFirst = True			' -- 처음 서식코드 저장시점 
    blnFile = False			' -- 파일저장유무 
    
    PrintLog "MakeHTF_W1101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1101MA1"

	Set lgcTB_3_2 = New C_TB_3_2		' -- 해당서식 클래스 
	
	If Not lgcTB_3_2.LoadData Then Exit Function
	
	Set cDataExists = new TYPE_DATA_EXIST_W1101MA1			
	
	Call lgcTB_3_2.Clone(oRs2)
	'==========================================
	' -- 제3호의2 표준대차대조표 오류검증 

	' -- 일반법인 - 자산 
	If lgcTB_3_2.W1 = "1" Then ' -- 일반법인 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_1) = "83" : arrHTFBody(TYPE_3) = "83"
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("A112", 4)		' 일반법인_자산 서식코드 
		arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNIChar("A172", 4)		' 일반법인_자산 서식코드 (합계)
		
		lgcTB_3_2.Filter "W2 LIKE '1%'"					' 레코드셋 필터 
		oRs2.Filter  = "W2 LIKE '1%'"	
	
		Do Until lgcTB_3_2.EOF 
		
		'Response.End 'ZZZ
		  SELECt Case  lgcTB_3_2.W4 
		     Case "01"
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '16'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'코드(01)유동자산= 코드 02 + 16
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(01)유동자산","코드 02 + 16"))
						   blnError = True	
					End If
					
			  Case "02"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '03'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'코드(02)당좌자산  = 코드 03 + 04 + 05 + 06 + 07 + 10 + 14 + 15	
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(02)당좌자산","코드 03 + 04 + 05 + 06 + 07 + 10 + 14 + 15	"))
						   blnError = True	
					End If		
			Case "07"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '08'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- 코드(07)단기대여금  = 코드 08 + 09     
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(07)단기대여금 ","코드 08 + 09 "))
						   blnError = True	
					End If	
			Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(10)미수금  = 코드 11 + 12 + 13 
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(07)단기대여금 ","코드 08 + 09 "))
						   blnError = True	
					End If	
			  Case "16"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '17'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '29'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'코드(16)재고자산  = 코드 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(16)재고자산","코드 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28 + 29"))
						   blnError = True	
					End If			
				Case "36"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '37'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(36)고정자산  = 코드 37 + 50 + 60
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(36)고정자산 ","코드 37 + 50 + 60"))
						   blnError = True	
					End If											
				
				Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '45'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(37)투자자산= 코드 38 + 39 + 40 + 44 + 45 + 46 + 47 + 48 + 49
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(37)투자자산","코드 38 + 39 + 40 + 44 + 45 + 46 + 47 + 48 + 49"))
						   blnError = True	
					End If		
					
					
		
			 
			    
			    
				Case "50"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '51'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- 코드(50)유형자산= 코드 51 + 52 + 53 + 54 + 55 + 56 + 57 + 58 + 59
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(50)유형자산","코드 51 + 52 + 53 + 54 + 55 + 56 + 57 + 58 + 59"))
						   blnError = True	
					End If		
					
				Case "60"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '61'"	
					dblAmt1 =UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '68'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '70'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '71'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'-  코드(60)무형자산= 코드 61 + 62 + 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70 + 71
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(60)무형자산","코드 61 + 62 + 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70 + 71"))
						   blnError = True	
					End If					
				Case "72"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '36'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					dblAmt72 =UNICDbl(lgcTB_3_2.DR_INV, 0)	
					
					'-  -- 코드(72)자산총계= 코드 01 + 36
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1 Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(72)자산총계","코드 01 + 36"))
						   blnError = True	
					End If	
					
												
			End Select
			 
			    ' -- 잔액 200703 삭제?
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.DR_INV, 16, 0)
			
				' -- 합계(차변)	
				If Not ChkNotNull(lgcTB_3_2.W5, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNINumeric(lgcTB_3_2.W5, 20, 0)
				' -- 합계(대변)
				If Not ChkNotNull(lgcTB_3_2.W6, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNINumeric(lgcTB_3_2.W6, 20, 0)
			
			lgcTB_3_2.MoveNext 
		Loop

		lgcTB_3_2.Filter ""			' -- 필터 해제 
		oRs2.Filter = ""
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("", 38)	' -- 공란 
		arrHTFBody(TYPE_3) = arrHTFBody(TYPE_3) & UNIChar("", 54)	' -- 공란 (합계)

		'----------------------------------------------------------------------
		'부채와 자본총계 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_2) = "83" : arrHTFBody(TYPE_4) = "84"
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("A113", 4)		' 일반법인_부채자본 서식코드 
		arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNIChar("A172", 4)		' 일반법인_자산 서식코드 (합계)

		lgcTB_3_2.Filter "W2 LIKE '2%' OR W2 LIKE '3%' OR W2 = '4'"					' 레코드셋 필터 
	
		oRs2.Filter = "W2 LIKE '2%' OR W2 LIKE '3%' OR W2 = '4'"					' 레코드셋 필터 
		Do Until lgcTB_3_2.EOF 
	
	    	Select Case  lgcTB_3_2.W4 
		        '코드(01)유동부채= 코드 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14
		        Case "01"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '02'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '03'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '04'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '06'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '11'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					''코드(01)유동부채= 코드 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(01)유동부채 ","코드 02 + 03 + 04 + 05 + 06 + 10 + 11 + 12 + 13 + 14"))
						   blnError = True	
					End If	
					
					
					
				 
		        Case "06"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '07'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					''코드(06)선수금  = 코드 07 + 08 + 09
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(06)선수금","코드 07 + 08 + 09"))
						   blnError = True	
					End If	
						
						
						
				  
		        Case "15"
		         
			        oRs2.MoveFirst
					oRs2.Find "W4 = '16'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '67'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '24'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					' '- 코드(15)고정부채= 코드 16 + 67 + 17 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(15)고정부채","코드 16 + 67 + 17 + 21 + 22 + 23 + 24 + 25 + 26 + 27 + 28"))
						   blnError = True	
					End If	
					
					
					
			  Case "17"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '18'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					
					'- 코드(17)장기차입금= 코드 18 + 19 + 20
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(17)장기차입금","코드 18 + 19 + 20"))
						   blnError = True	
					End If	
					
			
				
			    Case "29"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'- 코드(29)부채총계= 코드 01 + 15	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(29)부채총계","코드 01 + 15"))
						   blnError = True	
					End If	
					
					
				Case "41"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '42'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '43'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'-- 코드(41)자본금= 코드 42 + 43
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(41)자본금","코드 42 + 43"))
						   blnError = True	
					End If	
						
			
		   	   '코드(41)의 자본금이 0이 아니면 자본금과적립금조정명세서(갑) (A144)의  코드(01)자본금의 항목(5)기말잔액과 일치 
			
			  
			        If  UNICDbl(lgcTB_3_2.CR_INV, 0) <> 0 Then
			        	Set cDataExists.A144 = new C_TB_50A	' -- W7105MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A144.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A144.WHERE_SQL = " AND W_CD = '01' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A144.LoadData() Then
								blnError = True
						
							Else
							
							     sTmp =  UNICDbl(cDataExists.A144.W5,0)
							    
								If UNICDbl(lgcTB_3_2.CR_INV, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(41)의 자본금"," 자본금과적립금조정명세서(갑) (A144)의  코드(01)자본금의 항목(5)기말잔액"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A144 = Nothing				
					 End If		
			    
			    
			    
			       Case "44"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '45'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '48'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'-코드(44)자본잉여금= 코드 45 + 46 + 47 + 48 + 49
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(44)자본잉여금","코드 45 + 46 + 47 + 48 + 49"))
						   blnError = True	
					End If	
					
					
				 Case "50"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '51'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '55'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					'- 코드(50)이익잉여금= 코드 51 + 52 + 53 + 54 + 55 + 56
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(50)이익잉여금=","코드 51 + 52 + 53 + 54 + 55 + 56"))
						   blnError = True	
					End If	
	
			     Case "57"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '58'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '64'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					'- 코드(57)자본조정= 코드 58 + 59 + 60 + 61 + 62 + 63 + 64
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(57)자본조정=","코드 58 + 59 + 60 + 61 + 62 + 63 + 64"))
						   blnError = True	
					End If	
				
				
				Case "65"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '41'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '44'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					'	- 코드(65)자본총계= 코드 41 + 44 + 50 + 57
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(65)자본총계=","코드 41 + 44 + 50 + 57"))
						   blnError = True	
					End If	
				Case "66"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					dblAmt66 =  UNICDbl(lgcTB_3_2.CR_INV, 0)
					'	- 코드(66)부채와 자본총계= 코드 29 + 65
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(66)부채와 자본총계","코드 29 + 65"))
						blnError = True	
					End If		
					
					
			    End Select
			    
			' -- 잔액 
			If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.CR_INV, 16, 0)
				
			' -- 합계(차변)	
			If Not ChkNotNull(lgcTB_3_2.W5, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNINumeric(lgcTB_3_2.W5, 20, 0)
			' -- 합계(대변)
			If Not ChkNotNull(lgcTB_3_2.W6, lgcTB_3_2.W3) Then blnError = True	
			arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNINumeric(lgcTB_3_2.W6, 20, 0)
			
			lgcTB_3_2.MoveNext 
		Loop
		
			If dblAmt72 <> dblAmt66  Then
				Call SaveHTFError(lgsPGM_ID, dblAmt72, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(72)자산총계(" & dblAmt72 &")","코드(66)부채와 자본총계(" & dblAmt66 &")"))
				blnError = True	
		   End If	
		
		
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("", 58)	' -- 공란 
		arrHTFBody(TYPE_4) = arrHTFBody(TYPE_4) & UNIChar("", 44)	' -- 공란 (합계)
	
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_1)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_2)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_3)
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_4)
		
		' -- 파일에 기록한다.
		If Not blnError Then
			Call WriteLine2File(arrHTFBody(TYPE_1))
			Call WriteLine2File(arrHTFBody(TYPE_2))
			Call WriteLine2File(arrHTFBody(TYPE_3))
			Call WriteLine2File(arrHTFBody(TYPE_4))
			
			'Call PushRememberDoc(arrHTFBody(TYPE_3))	' -- 바로 출력하지 않고 기억시킨다(inc_HomeTaxFunc.asp에 정의)
			'Call PushRememberDoc(arrHTFBody(TYPE_4))
		End If		

	
	Else	' -- 금융법인 
		'----------------------------------------------------------------------
		arrHTFBody(TYPE_1) = "83" : arrHTFBody(TYPE_2) = "85"
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("A114", 4)		' 금융법인 서식코드 
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("A172", 4)		' 금융법인 서식코드 
		
		Do Until lgcTB_3_2.EOF 
		
			If Left(lgcTB_3_2.W1, 1) = "1" Then
			
			   Select Case lgcTB_3_2.W4 	
				Case "01"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- 코드(01)현금과예치금= 코드 02 + 03 + 04
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(01)현금과예치금","코드 02 + 03 + 04"))
						blnError = True	
					End If	
				Case "05"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '06'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '07'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '08'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '09'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- 코드(05)상품유가증권= 코드 06 + 07 + 08 + 09
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(05)상품유가증권","코드 06 + 07 + 08 + 09"))
						blnError = True	
					End If				
					
				Case "10"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '11'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '12'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '13'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '14'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- 코드(10)투자유가증권= 코드 11 + 12 + 13 + 14
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(10)투자유가증권","코드 11 + 12 + 13 + 14"))
						blnError = True	
					End If			
				Case "15"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '16'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '17'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '18'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '19'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '20'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '21'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(15)대출채권= 코드 16 + 17 + 18 + 19 + 20 + 21
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(15)대출채권","코드 16 + 17 + 18 + 19 + 20 + 21"))
						blnError = True	
					End If
				Case "23"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '24'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '25'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '26'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '27'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(23)리스자산= 코드 24 + 25 + 26 + 27
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(23)리스자산","코드 24 + 25 + 26 + 27"))
						blnError = True	
					End If	
				Case "28"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '29'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '30'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '35'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '36'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- - 코드(28)고정자산= 코드 29 + 30 + 35 + 36
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(28)고정자산"," 코드 29 + 30 + 35 + 36"))
						blnError = True	
					End If	
				Case "30"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '31'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '32'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '33'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '34'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					'- 코드(30)유형자산= 코드 31 + 32 + 33 + 34
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(30)유형자산"," 코드 31 + 32 + 33 + 34"))
						blnError = True	
					End If		
																			
				Case "37"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '38'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '39'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '40'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '41'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '42'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					'- - 코드(37)기타자산= 코드 38 + 39 + 40 + 41 + 42
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(37)기타자산"," 코드 38 + 39 + 40 + 41 + 42"))
						blnError = True	
					End If			
				
				Case "43"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '01'"	
					dblAmt1 = UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '05'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '10'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '15'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '22'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '23'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '28'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					oRs2.Find "W4 = '37'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("DR_INV"), 0)
					
					
					dblAmt43 =  UNICDbl(lgcTB_3_2.DR_INV, 0)
					'코드(43)자산총계= 코드 01 + 05 + 10 + 15 + 22 + 23 + 28 + 37
					If UNICDbl(lgcTB_3_2.DR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.DR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(43)자산총계"," 코드01 + 05 + 10 + 15 + 22 + 23 + 28 + 37"))
						blnError = True	
					End If				
				 End Select
				
			
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.W3, DR_INV, 0)
				
				If Not ChkNotNull(lgcTB_3_2.DR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.W3, DR_INV, 0)
			Else
			
			   Select Case lgcTB_3_2.W4 	
			   	Case "44"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '45'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '46'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
					
					'- 코드(44)신용부채= 코드 45 + 46
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(44)신용부채"," 코드 45 + 46"))
						blnError = True	
					End If		
																											
				
				
				Case "47"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '48'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '49'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '50'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '51'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '52'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
										
					'- 코드(47)차입금= 코드 48 + 49 + 50 + 51 + 52	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(47)차입금"," 코드 48 + 49 + 50 + 51 + 52"))
						blnError = True	
					End If																							
			
				Case "54"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '55'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '56'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '57'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '58'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '59'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '60'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '61'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '62'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
										
					'- 코드(54)기타부채= 코드 55 + 56 + 57 + 58 + 59 + 60 + 61 + 62
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(54)기타부채"," 코드  55 + 56 + 57 + 58 + 59 + 60 + 61 + 62"))
						blnError = True	
					End If		
					
					
					
				Case "63"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '64'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '65'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '66'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- 코드(63)보험사 제준비금= 코드 64 + 65 + 66	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(63)보험사 제준비금"," 코드 64 + 65 + 66"))
						blnError = True	
					End If					
				
				Case "67"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '44'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '47'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '53'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '54'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '63'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- 코드(67)부채총계= 코드 44 + 47 + 53 + 54 + 63
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(67)부채총계"," 코드 44 + 47 + 53 + 54 + 63"))
						blnError = True	
					End If		
				
				Case "68"
			     
						
			
		   	   '코드(68)의 자본금이 0이 아니면 자본금과적립금조정명세서(갑) (A144)의  코드(01)자본금의 항목(5)기말잔액과 일치 
			
			  
			        If  UNICDbl(lgcTB_3_2.CR_INV, 0) <> 0 Then
			        	Set cDataExists.A144 = new C_TB_50A	' -- W7105MA1_HTF.asp 에 정의됨 
								
							' -- 추가 조회조건을 읽어온다.
							cDataExists.A144.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
							cDataExists.A144.WHERE_SQL = " AND W_CD = '01' "			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
								
							If Not cDataExists.A144.LoadData() Then
								blnError = True
						
							Else
							
							     sTmp =  UNICDbl(cDataExists.A144.W5,0)
							    
								If UNICDbl(lgcTB_3_2.CR_INV, 0) <> UNICDbl(sTmp , 0) Then
									blnError = True
									Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(68)의 자본금"," 자본금과적립금조정명세서(갑) (A144)의  코드(01)자본금의 항목(5)기말잔액"))
								End If
							End If
					
							
							' -- 사용한 클래스 메모리 해제 
							Set cDataExists.A144 = Nothing				
					 End If		
			    
			    
			    	
				
				
				Case "69"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '70'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '71'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '72'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- 코드(69)자본잉여금= 코드 70 + 71 + 72
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(69)자본잉여금"," 코드 70 + 71 + 72"))
						blnError = True	
					End If		
				
				Case "73"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '74'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '75'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '76'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '77'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '78'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					'- - 코드(73)이익잉여금= 코드 74 + 75 + 76 + 77 + 78
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(73)이익잉여금"," 코드 74 + 75 + 76 + 77 + 78"))
						blnError = True	
					End If	
				
				Case "79"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '80'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '81'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '82'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)

															
					' 코드(79)자본조정= 코드 80 + 81 + 82	
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(79)자본조정"," 코드  80 + 81 + 82	"))
						blnError = True	
					End If		
						
				Case "83"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '68'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '69'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '73'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '79'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
															
					' - 코드(83)자본총계= 코드 68 + 69 + 73 + 79
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(83)자본총계"," 코드  68 + 69 + 73 + 79"))
						blnError = True	
					End If				
				
				Case "84"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '67'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '83'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					
															
					' - 코드(84)부채 및 자본총계= 코드 67 + 83
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(84)부채 및 자본총계"," 코드  67 + 83"))
						blnError = True	
					End If		
				
				Case "84"
			        oRs2.MoveFirst
					oRs2.Find "W4 = '67'"	
					dblAmt1 = UNICDbl(oRs2("CR_INV"), 0)
					oRs2.Find "W4 = '83'"	
					dblAmt1 = dblAmt1 + UNICDbl(oRs2("CR_INV"), 0)
					dblAmt84 = UNICDbl(lgcTB_3_2.CR_INV, 0)
															
					' - 코드(84)부채 및 자본총계= 코드 67 + 83
					If UNICDbl(lgcTB_3_2.CR_INV, 0) <> dblAmt1  Then
						Call SaveHTFError(lgsPGM_ID, lgcTB_3_2.CR_INV, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(84)부채 및 자본총계"," 코드  67 + 83"))
						blnError = True	
					End If	
						
			    End Select
			    
				If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNINumeric(lgcTB_3_2.W3, CR_INV, 0)
				
				If Not ChkNotNull(lgcTB_3_2.CR_INV, lgcTB_3_2.W3) Then blnError = True	
				arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNINumeric(lgcTB_3_2.W3, CR_INV, 0)
			End If	
		
			lgcTB_3_2.MoveNext 
		Loop
		
		If dblAmt43 <> dblAmt84  Then
			Call SaveHTFError(lgsPGM_ID, dblAmt43, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(43)자산총계(" & dblAmt43 &")","코드(84)부채와 자본총계(" & dblAmt84 &")"))
			blnError = True	
		 End If
		
		arrHTFBody(TYPE_1) = arrHTFBody(TYPE_1) & UNIChar("", 50)	' -- 공란 
		arrHTFBody(TYPE_2) = arrHTFBody(TYPE_2) & UNIChar("", 134)	' -- 공란 
	
		PrintLog "WriteLine2File : " & arrHTFBody(TYPE_1)
		
		' -- 파일에 기록한다.
		If Not blnError Then
			Call WriteLine2File(arrHTFBody(TYPE_1))
			Call WriteLine2File(arrHTFBody(TYPE_2))
			'Call PushRememberDoc(arrHTFBody(TYPE_2))	' -- 바로 출력하지 않고 기억시킨다(inc_HomeTaxFunc.asp에 정의)
		End If
	End If
					
	' ----------- 
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_3_2 = Nothing	' -- 메모리해제 
	
End Function


%>
