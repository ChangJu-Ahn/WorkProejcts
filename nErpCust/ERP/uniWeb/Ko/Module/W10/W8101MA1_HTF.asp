
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제3호 법인세 과세표준 및 세액조정계산서 
'*  3. Program ID           : W8101MA1
'*  4. Program Name         : W8101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_3

Set lgcTB_3 = Nothing ' -- 초기화 

Class C_TB_3
	' -- 테이블의 컬럼변수 
	Dim W01
	Dim W02
	Dim W03
	Dim W04
	Dim W05
	Dim W54
	Dim W06
	Dim W07
	Dim W08
	Dim W09
	Dim W10
	Dim W11
	Dim W12
	Dim W13
	Dim W14_CD
	Dim W14
	Dim W15
	Dim W16
	Dim W17
	Dim W18
	Dim W19
	Dim W20
	Dim W21
	Dim W22
	Dim W23
	Dim W24
	Dim W25_NM
	Dim W25
	Dim W26
	Dim W27
	Dim W28
	Dim W29
	Dim W30
	Dim W31
	Dim W32
	Dim W33
	Dim W34
	Dim W35_CD
	Dim W35
	Dim W36
	Dim W37
	Dim W38
	Dim W39
	Dim W40
	Dim W41
	Dim W42
	Dim W43_NM
	Dim W43
	Dim W44
	Dim W45
	Dim W46
	Dim W47
	Dim W48
	Dim W49
	Dim W50
	Dim W51
	Dim W52
	Dim W53
	Dim W55
	
	' -- 2005-01-04 : 200603 개정 
	Dim W55_1
	Dim W56

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

		' -- 제1호서식을 읽어온다.
		Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs1,lgStrSQL, "", "") = False Then
		    Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
		    Exit Function
		End If

		W01			= oRs1("W01")
		W02			= oRs1("W02")
		W03			= oRs1("W03")
		W04			= oRs1("W04")
		W05			= oRs1("W05")
		W54			= oRs1("W54")
		W06			= oRs1("W06")
		W07			= oRs1("W07")
		W08			= oRs1("W08")
		W09			= oRs1("W09")
		W10			= oRs1("W10")
		W11			= oRs1("W11")
		W12			= oRs1("W12")
		W13			= oRs1("W13")
		W14_CD			= oRs1("W14_CD")
		W14			= oRs1("W14")
		W15			= oRs1("W15")
		W16			= oRs1("W16")
		W17			= oRs1("W17")
		W18			= oRs1("W18")
		W19			= oRs1("W19")
		W20			= oRs1("W20")
		W21			= oRs1("W21")
		W22			= oRs1("W22")
		W23			= oRs1("W23")
		W24			= oRs1("W24")
		W25_NM			= oRs1("W25_NM")
		W25			= oRs1("W25")
		W26			= oRs1("W26")
		W27			= oRs1("W27")
		W28			= oRs1("W28")
		W29			= oRs1("W29")
		W30			= oRs1("W30")
		W31			= oRs1("W31")
		W32			= oRs1("W32")
		W33			= oRs1("W33")
		W34			= oRs1("W34")
		W35_CD			= oRs1("W35_CD")
		W35			= oRs1("W35")
		W36			= oRs1("W36")
		W37			= oRs1("W37")
		W38			= oRs1("W38")
		W39			= oRs1("W39")
		W40			= oRs1("W40")
		W41			= oRs1("W41")
		W42			= oRs1("W42")
		W43_NM			= oRs1("W43_NM")
		W43			= oRs1("W43")
		W44			= oRs1("W44")
		W45			= oRs1("W45")
		W46			= oRs1("W46")
		W47			= oRs1("W47")
		W48			= oRs1("W48")
		W49			= oRs1("W49")
		W50			= oRs1("W50")
		W51			= oRs1("W51")
		W52			= oRs1("W52")
		W53			= oRs1("W53")
		W55			= oRs1("W55")
		
		' -- 2005-01-04 : 200603 개정 
		W55_1		= oRs1("W55_1")
		W56			= oRs1("W56")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_3	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
End Class

' -- 본 서식에서 참고할 다른 서식들 
Class TYPE_DATA_EXIST_W8101MA1
	Dim A115
	Dim A116
	Dim A102
	Dim A144
	Dim A106
	Dim A159
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8101MA1"

	Set lgcTB_3 = New C_TB_3		' -- 해당서식 클래스 
	
	If Not lgcTB_3.LoadData Then Exit Function			' -- 제3호 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W8101MA1

	' -- 쿼리변수 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

	'==========================================
	' -- 제3호 법인세과세표준 및 세액조정계산서 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If ChkNotNull(lgcTB_3.W01, "결산서상당기순손익") Then ' -- 데이타존재시 검증식 
		If lgcTB_1.W2 <> "50" Then ' -- 법인종류별구분이 당기순이익과세법인 '50' 이 아닌 경우 
			
			' -- 제3호3(1)(2)표준손익계산서(A115)서식 
			Set cDataExists.A115 = new C_TB_3_3	' -- W1105MA1_HTF.asp 에 정의됨 
			
			' -- 추가 조회조건을 읽어온다.
			Call SubMakeSQLStatements_W8101MA1("A115_1",iKey1, iKey2, iKey3)   
			
			cDataExists.A115.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
			cDataExists.A115.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
			If Not cDataExists.A115.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제3호3(1)(2)표준손익계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
				'표준손익계산서(A115,A116)의 당기순손익(일반법인은 코드(82) 금융.보험.증권업법인은 코드(73))과일치하지 않으면 오류 
				If UNICDbl(lgcTB_3.W01, 0) <> UNICDbl(cDataExists.A115.W5, 0) Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, lgcTB_3.W01, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "결산서상당기순손익","표준손익계산서(A115,A116)의 당기순손익"))
				End If
			End If
		
			' -- 사용한 클래스 메모리 해제 
			Set cDataExists.A115 = Nothing
		End If
	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W01, 15, 0)
	
	If ChkNotNull(lgcTB_3.W02, "소득조정금액_익금산입") Then ' -- 데이타존재시 검증식 
		' -- 제15호 과목별소득금액조정명세서(A102)서식 
		Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp 에 정의됨 
			
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W8101MA1("A102_1",iKey1, iKey2, iKey3)   
			
		cDataExists.A102.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A102.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
		If Not cDataExists.A102.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "제15호 과목별소득금액조정명세서_익금산입및손금불산입", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
			''소득금액조정합계표(A102)의 익금산입및손금불산입의 금액(2)의 합계와 일치 
			If UNICDbl(lgcTB_3.W02, 0) <> UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소득조정금액_익금산입","소득금액조정합계표(A102)의 익금산입및손금불산입의 금액(2)의 합계"))
			End If
		End If
		' -- 사용한 클래스 메모리 해제 
		'Set cDataExists.A102 = Nothing

		' -- 제3호3(1)(2)표준손익계산서(A115)서식 
		Set cDataExists.A115 = new C_TB_3_3	' -- W1105MA1_HTF.asp 에 정의됨 
			
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W8101MA1("A115_2",iKey1, iKey2, iKey3)   
			
		cDataExists.A115.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A115.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
		
		If Not cDataExists.A115.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "제3호3(1)(2)표준손익계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
 
			'표준손익계산서(A115,A116)의 항목(81)의 법인세비용(일반법인)/항목(72)법인세비용(금융법인)보다 작으면 오류 
			If UNICDbl(lgcTB_3.W02, 0)< UNICDbl(cDataExists.A115.W5, 0) Then

				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg(TYPE_CHK_LOW_AMT, "소득조정금액_익금산입","표준손익계산서(A115,A116)의 항목(81)의 법인세비용(일반법인)/항목(72)법인세비용(금융법인)"))
			End If
			
		End If	
		
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A115 = Nothing
					
	Else
		blnError = True
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W02, 15, 0)
	
	If  ChkNotNull(lgcTB_3.W03, "소득조정금액_손금산입") Then 
		' -- 제15호 과목별소득금액조정명세서(A102)서식 
		Set cDataExists.A102 = new C_TB_15	' -- W5103MA1_HTF.asp 에 정의됨 
			
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W8101MA1("A102_2",iKey1, iKey2, iKey3)   
			
		cDataExists.A102.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A102.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
		If Not cDataExists.A102.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "제15호 과목별소득금액조정명세서_손금산입및익금불산입", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
			'소득금액조정합계표(A102)의 손금산입및익금불산입의 금액(5)의 합계와 일치 
			If UNICDbl(lgcTB_3.W03, 0) <> UNICDbl(cDataExists.A102.GetData("W2"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W03, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "소득조정금액_손금산입","소득금액조정합계표(A102)의 손금산입및익금불산입의 금액(5)의 합계"))
			End If
		End If
		' -- 사용한 클래스 메모리 해제 
		'Set cDataExists.A102 = Nothing

	Else
		blnError = True
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W03, 15, 0)
		
	If Not ChkNotNull(lgcTB_3.W04, "차가감소득금액") Then blnError = True	' -- 프로그램에서검증했으므로 패스 
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W04, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W05, "기부금한도초과액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W05, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W54, "기부금한도초과 이월액손금산입") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W54, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W06, "각사업연도 소득금액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W06, 15, 0)

	If Not ChkNotNull(lgcTB_3.W07, "이월결손금") Then blnError = True
	If Not ChkMinusAmt(lgcTB_3.W07, "이월결손금") Then blnError = True	' -- 음수체크 
	If UNICDbl(lgcTB_3.W06, 0) < 0 And UNICDbl(lgcTB_3.W07, 0) > 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, "코드(06)각사업년도소득금액이 음수인데 금액이 있으면 오류")
	End If
	If UNICDbl(lgcTB_3.W07, 0) > UNICDbl(lgcTB_3.W06, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, UNIGetMesg(TYPE_CHK_LOW_AMT, "각사업연도 소득금액","이월결손금"))
	End If
	
	If UNICDbl(lgcTB_3.W07, 0) > 0 Then
		' -- 0보다 큰 경우 A144 반드시 존재 
		
		' -- 이월결손금 등록 프로그램 
		Set cDataExists.A144 = new C_W7107MA1	' -- W7107MA1_HTF.asp 에 정의됨: 프로그램이라 클래스명이 프로그램ID
			
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W8101MA1("A144",iKey1, iKey2, iKey3)   
			
		cDataExists.A144.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A144.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
		If Not cDataExists.A144.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "이월결손금 등록", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
			'이월결손금 등록 (6)당기공제의 합계와 일치 
			If UNICDbl(lgcTB_3.W07, 0) <> UNICDbl(cDataExists.A144.W6, 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W07, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "이월결손금","이월결손금 등록_(6)당기공제의 합계"))
			End If
		End If	
		
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A144 = Nothing
	End If

	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W07, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W08, "비과세소득") Then blnError = True
	If Not ChkMinusAmt(lgcTB_3.W08, "비과세소득") Then blnError = True	' -- 음수체크 
	
	If UNICDbl(lgcTB_3.W06, 0) < 0 And UNICDbl(lgcTB_3.W08, 0) > 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W08, "코드(06)각사업년도소득금액이 음수인데 금액이 있으면 오류")
	End If
	If UNICDbl(lgcTB_3.W08, 0) > (UNICDbl(lgcTB_3.W06, 0) - UNICDbl(lgcTB_3.W07, 0)) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W08, UNIGetMesg(TYPE_CHK_HIGH_AMT, "비과세소득","코드(06)각사업년도소득금액 - 코드(07)이월결손금"))
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W08, 15, 0)
	
	
	
	If Not ChkNotNull(lgcTB_3.W09, "소득공제") Then blnError = True
	
	'2007.03.27 추가 lws
	'소득공제 0보다 크면 
	' <= 항목(108)각사업연도소득금액-이월결손금(코드07)-비과세소득(코드08) 
	
	if lgcTB_3.W09 > 0 then
	
		if lgcTB_3.W09  <= lgcTB_3.W06 - lgcTB_3.W07 - lgcTB_3.W08 then 'pass
		else
			Call SaveHTFError(lgsPGM_ID, lgcTB_3.W09, UNIGetMesg(TYPE_CHK_HIGH_AMT, "소득공제","항목(108)각사업연도소득금액-이월결손금(코드07)-비과세소득(코드08) "))
		end if

	end if
	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W09, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W10, "과세표준금액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W10, 15, 0)
	
	If Not ChkNotNull(UNICDBl(lgcTB_3.W11,0) * 100, "세율") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(UNICDBl(lgcTB_3.W11,0) * 100, 5, 2)

	If Not ChkNotNull(lgcTB_3.W12, "산출세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W12, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W13, "지점유보소득") Then blnError = True
	If UNICDbl(lgcTB_3.W13, 0) <> 0 And lgcTB_1.W1 = "2" Then
		If Not SearchTaxDocCd("A162") Then	' wa101mb2.asp 에 정의됨 
			blnError = True
			Call SaveHTFError(lgsPGM_ID, lgcTB_3.W13, "외국법인이고 지점보유소득이 '0'이 아닌경우 지점유보소득금액계산서(A162) 오류")
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W13, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W14, "세율") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W14, 5, 2)
	
	If Not ChkNotNull(lgcTB_3.W15, "산출세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W15, 15, 0)
 
	If Not ChkNotNull(lgcTB_3.W16, "산출세액합계") Then blnError = True
	If UNICDbl(lgcTB_3.W16, 0) <> UNICDbl(lgcTB_3.W15, 0)+ UNICDbl(lgcTB_3.W12, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W16, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "산출세액합계","산출세액"))
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W16, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W17, "공제감면세액(ㄱ)") Then blnError = True
	' 0 보다 큰 경우 A106 체크 (8호갑)
	If UNICDbl(lgcTB_3.W17, 0) > 0 Then
		' -- 제8호(갑)공제감면...
		
		' --- 데이타 조회 SQL
		Call SubMakeSQLStatements_W8101MA1("A106",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,oRs2,lgStrSQL, "", "") = False Then
			blnError = True
		    Call SaveHTFError(lgsPGM_ID, "제8호(갑)공제감면세액및추가납부세액합계표", TYPE_DATA_NOT_FOUND)
		Else
		
			If UNICDbl(lgcTB_3.W17, 0) <> UNICDbl(oRs2("W20"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W17, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제감면세액(ㄱ)","제8호(갑)공제감면세액및추가납부세액합계표"))
			End If
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W17, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W18, "차감세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W18, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W19, "공제감면세액(ㄴ)") Then blnError = True
	If UNICDbl(lgcTB_3.W19, 0) > UNICDbl(lgcTB_3.W18, 0) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W19, UNIGetMesg(TYPE_CHK_HIGH_AMT, "공제감면세액(ㄴ)","차감세액"))
	End If	
	If UNICDbl(lgcTB_3.W19, 0) > 0 Then
		If Not oRs2 is Nothing Then	' -- 위에서 로드되었다면..
			If UNICDbl(lgcTB_3.W19, 0) <> UNICDbl(oRs2("W21"), 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W19, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "공제감면세액(ㄴ)","제8호(갑)공제감면세액및추가납부세액합계표"))
			End If
		End If
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W19, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W20, "가산세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W20, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W21, "가감계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W21, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W22, "기한내납부세액_중간예납세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W22, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W23, "기한내납부세액_수시부과세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W23, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W24, "기한내납부세액_원천납부세액") Then blnError = True

	' 2006-01-04 : 200603 개정 
	'If UNICDbl(lgcTB_3.W24, 0) <> 0 Then
		' -- 제10호 
		Set cDataExists.A159 = new C_TB_10A	' -- W7101MA1_HTF.asp 에 정의됨 
			
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W8101MA1("A159",iKey1, iKey2, iKey3)   
			
		cDataExists.A159.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A159.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
			
		If Not cDataExists.A159.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "제10호 원천납부세액명세서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
			'제10호 원천납부세액명세서 (6)법인세합계와 일치 
			If UNICDbl(lgcTB_3.W24, 0) <> UNICDbl(cDataExists.A159.W6, 0) Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, lgcTB_3.W24, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기한내납부세액_원천납부세액","제10호 원천납부세액명세서_(6)법인세합계"))
			End If
		End If	
		
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A159 = Nothing
	'End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W24, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W25, "기한내납부세액_간접투자회사등의 외국납부세액") Then blnError = True
	
	If UNICDbl(lgcTB_3.W25, 0) <> 0 Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W25, UNIGetMesg("코드(25)의 기한내납부세액_간접투자회사등의 외국납부세액이 0보다 큰 경우 간접투자등의 외국납부세액 계산서(별지 제10호의 2 서식)가 있어야합니다. 해당서식이 존재하지 않으니 uniERP 팀에 문의하시기 바랍니다", "",""))
	End If
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W25, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W26, "기한내납부세액_소계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W26, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W27, "신고납부전가세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W27, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W28, "기납부세액합계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W28, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W29, "감면분추가납부세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W29, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W30, "납부할세액계산_차감납무세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W30, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W31, "토지등양도소득에대한법인세계산_양도차익_등기자산") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W31, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W32, "토지등양도소득에대한법인세계산_양도차익_미등기자산") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W32, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W33, "비과세소득") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W33, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W34, "과세표준") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W34, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W35, "세율") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W35, 5, 2)
	
	If Not ChkNotNull(lgcTB_3.W36, "산출세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W36, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W37, "감면세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W37, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W38, "차감세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W38, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W39, "공제세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W39, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W40, "가산세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W40, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W41, "가감계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W41, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W42, "기납부세액_수시부과세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W42, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W43, "기납부세액_(" & lgcTB_3.W43_NM & ")세액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W43, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W44, "기납부세액_계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W44, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W45, "차감납부할세액_토지등양도소득에대한법인세계산") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W45, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W46, "세액계_차감납부할세액계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W46, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W55, "세액계_사실과다른회계처리경정세액공제") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W55, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W47, "세액계_분납세액계산범위액") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W47, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W48, "분납할세액_현금납부") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W48, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W49, "분납할세액_물납") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W49, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W50, "분납할세액_계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W50, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W51, "차감납부세액_현금납부") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W51, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W52, "차감납부세액_물납") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W52, 15, 0)
	
	If Not ChkNotNull(lgcTB_3.W53, "차감납부세액_계") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W53, 15, 0)

	' -- 2006-01-04 : 200603 개정 
	If Not ChkNotNull(lgcTB_3.W55_1, "선박표준이익") Then blnError = True
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W55_1, 15, 0)

	If UNICDbl(lgcTB_3.W55_1, 0) > 0 Then	' -- 선박표준이익이 0보다 큰 경우 선박표준이익 산출명세서(A224)의 항목(7)선박표준이익 금액과 일치 비교(미개발)
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W55_1, "선박표준이익이 0보다 큰 경우 선박표준이익 산출명세서(A224)의 항목(7)선박표준이익 금액과 일치 비교")
	End If
	
	If UNICDbl(lgcTB_3.W56, 0) <> (UNICDbl(lgcTB_3.W10, 0) + UNICDbl(lgcTB_3.W55_1, 0)) Then
		blnError = True
		Call SaveHTFError(lgsPGM_ID, lgcTB_3.W56, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(56)과세표준금액","코드(10)과세표준금액 + 코드(55)선박표준이익"))
	End If	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_3.W56, 15, 0)

	'sHTFBody = sHTFBody & UNIChar("", 69)	' -- 공란 
	sHTFBody = sHTFBody & UNIChar("", 19)	' -- 공란 : 2006-01-05 : 200603 개정판 
	
	' ----------- 
	Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	
 
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_3 = Nothing	' -- 메모리해제 

End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8101MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A115_1" '-- 외부 참조 SQL
			
			lgStrSQL = ""
			' -- 표준손익계산서(A115,A116)의 당기순손익(일반법인은 코드(82) 금융.보험.증권업법인은 코드(73))과 일치하지 않으면 오류 
			'lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- 표준손익계산서 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- 법인구분(일반/금융)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '82'"		 	 & vbCrLf	' -- 법인구분(일반)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '73'"		 	 & vbCrLf	' -- 법인구분(금융)
			End If

	  Case "A115_2" '-- 외부 참조 SQL
			
			lgStrSQL = ""
			' -- 표준손익계산서(A115,A116)의 항목(81)법인세비용(일반법인)/항목(72)법인세비용(금융법인) 
			'lgStrSQL = lgStrSQL & "	AND A.BS_PL_FG	= '2'"		 	 & vbCrLf	' -- 표준손익계산서 
			lgStrSQL = lgStrSQL & "	AND A.W1		= '" & lgcCompanyInfo.COMP_TYPE2 & "'"		 	 & vbCrLf	' -- 법인구분(일반/금융)
			If lgcCompanyInfo.COMP_TYPE2 = "1" Then
				lgStrSQL = lgStrSQL & "	AND A.W4		= '81'"		 	 & vbCrLf	' -- 법인구분(일반)
			Else
				lgStrSQL = lgStrSQL & "	AND A.W4		= '72'"		 	 & vbCrLf	' -- 법인구분(금융)
			End If
			
	  Case "A102_1" '-- 외부 참조 SQL
		
			' -- 소득금액조정합계표(A102)의 익금산입및손금불산입의 금액(2)의 합계와 일치 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '1'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf

	  Case "A102_2" '-- 외부 참조 SQL
	  
			lgStrSQL = ""
			' -- 소득금액조정합계표(A102)의 손금산입및익금불산입의 금액(5)의 합계와 일치 
			lgStrSQL = lgStrSQL & "	AND A.W_TYPE	= '2'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf
			
	  Case "A144" '-- 외부 참조 SQL
	  
			lgStrSQL = ""
			' -- 이월결손금등록 
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf	
	  
	  Case "A106"
			lgStrSQL = ""
			' -- 제8호(갑) 공제감면세액	
			lgStrSQL = lgStrSQL & "SELECT " & vbCrLf
			lgStrSQL = lgStrSQL & " ISNULL( ( " & vbCrLf
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf	
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '50' " & vbCrLf	
			lgStrSQL = lgStrSQL & "	), 0) - ISNULL( (" & vbCrLf	
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf	' 
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '66' " & vbCrLf
			lgStrSQL = lgStrSQL & "	), 0) AS W20 " & vbCrLf	
			lgStrSQL = lgStrSQL & ",ISNULL( ( " & vbCrLf
			lgStrSQL = lgStrSQL & "		SELECT W4 " & vbCrLf					' -- 그리드1 (10)소계 
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '0' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '10' " & vbCrLf	
			lgStrSQL = lgStrSQL & "	), 0) + ISNULL( (" & vbCrLf	
			lgStrSQL = lgStrSQL & "		SELECT W7 " & vbCrLf					' -- (66) 연구인력..(최저한세..)		
			lgStrSQL = lgStrSQL & "		FROM TB_8A A WITH (NOLOCK) " & vbCrLf	
			lgStrSQL = lgStrSQL & "		WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND A.W_TYPE = '2' " & vbCrLf	
			lgStrSQL = lgStrSQL & "			AND A.W2_1	= '66' " & vbCrLf
			lgStrSQL = lgStrSQL & "	), 0) AS W21 " & vbCrLf		
			lgStrSQL = lgStrSQL & " " & vbCrLf	

	  Case "A159" '-- 외부 참조 SQL
	  
			lgStrSQL = ""
			' -- 이월결손금등록 
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO	= 999999" 	 & vbCrLf	
	  
	End Select
	PrintLog "SubMakeSQLStatements_W8101MA1 : " & lgStrSQL
End Sub

%>
