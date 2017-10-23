<%
'======================================================================================================
'*  1. Function Name        : 신고구분 공통 
'*  3. Program ID           : WB101MA1
'*  4. Program Name         : WB101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcCompanyInfo

Set lgcCompanyInfo = Nothing	' -- 초기화 

Class C_COMPANY_HISTORY
	' -- 테이블의 컬럼변수 
	Dim TAX_DOC_CD
	Dim CO_NM
	Dim CO_ADDR
	Dim OWN_RGST_NO
	Dim LAW_RGST_NO
	Dim REPRE_NM
	Dim REPRE_RGST_NO
	Dim TEL_NO
	Dim COMP_TYPE1
	Dim DEBT_MULTIPLE
	Dim COMP_TYPE2
	Dim TAX_OFFICE
	Dim HOLDING_COMP_FLG
	Dim IND_CLASS
	Dim IND_TYPE
	Dim FOUNDATION_DT
	Dim HOME_TAX_USR_ID
	Dim HOME_TAX_E_MAIL
	Dim HOME_TAX_MAIN_IND
	Dim FISC_START_DT
	Dim FISC_END_DT
	Dim HOME_ANY_START_DT
	Dim HOME_ANY_END_DT
	Dim BANK_CD
	Dim BANK_BRANCH
	Dim BANK_DPST
	Dim BANK_ACCT_NO
	Dim INCOM_DT
	Dim HOME_FILE_MAKE_DT
	Dim REVISION_YM
	Dim EX_RECON_FLG
	Dim EX_54_FLG
	Dim AGENT_NM
	Dim RECON_BAN_NO
	Dim RECON_MGT_NO
	Dim AGENT_TEL_NO
	Dim AGENT_RGST_NO
	Dim REQUEST_DT
	Dim APPO_NO
	Dim APPO_DT
	Dim APPO_DESC
	Dim File_Name
	
	Dim lgoRs1	' -- 조정반그리 
	
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

		CO_NM			= oRs1("CO_NM")
		CO_ADDR			= oRs1("CO_ADDR")
		OWN_RGST_NO		= oRs1("OWN_RGST_NO")
		LAW_RGST_NO		= oRs1("LAW_RGST_NO")
		REPRE_NM			= oRs1("REPRE_NM")
		REPRE_RGST_NO	= oRs1("REPRE_RGST_NO")
		TEL_NO			= oRs1("TEL_NO")
		COMP_TYPE1		= oRs1("COMP_TYPE1")
		DEBT_MULTIPLE	= oRs1("DEBT_MULTIPLE")
		COMP_TYPE2		= oRs1("COMP_TYPE2")
		TAX_OFFICE		= oRs1("TAX_OFFICE")
		HOLDING_COMP_FLG	= oRs1("HOLDING_COMP_FLG")
		IND_CLASS		= oRs1("IND_CLASS")
		IND_TYPE			= oRs1("IND_TYPE")
		FOUNDATION_DT	= oRs1("FOUNDATION_DT")
		HOME_TAX_USR_ID	= oRs1("HOME_TAX_USR_ID")
		HOME_TAX_E_MAIL	= oRs1("HOME_TAX_E_MAIL")
		HOME_TAX_MAIN_IND= oRs1("HOME_TAX_MAIN_IND")
		FISC_START_DT	= oRs1("FISC_START_DT")
		FISC_END_DT		= oRs1("FISC_END_DT")
		HOME_ANY_START_DT= oRs1("HOME_ANY_START_DT")
		HOME_ANY_END_DT	= oRs1("HOME_ANY_END_DT")
		BANK_CD			= oRs1("BANK_CD")
		BANK_BRANCH		= oRs1("BANK_BRANCH")
		BANK_DPST		= oRs1("BANK_DPST")
		BANK_ACCT_NO		= oRs1("BANK_ACCT_NO")
		INCOM_DT			= oRs1("INCOM_DT")
		HOME_FILE_MAKE_DT= oRs1("HOME_FILE_MAKE_DT")
		REVISION_YM		= oRs1("REVISION_YM")
		EX_RECON_FLG		= oRs1("EX_RECON_FLG")
		EX_54_FLG		= oRs1("EX_54_FLG")
		AGENT_NM			= oRs1("AGENT_NM")
		RECON_BAN_NO		= oRs1("RECON_BAN_NO")
		RECON_MGT_NO		= oRs1("RECON_MGT_NO")
		AGENT_TEL_NO		= oRs1("AGENT_TEL_NO")
		AGENT_RGST_NO	= oRs1("AGENT_RGST_NO")
		REQUEST_DT		= oRs1("REQUEST_DT")
		APPO_NO			= oRs1("APPO_NO")
		APPO_DT			= oRs1("APPO_DT")
		APPO_DESC		= oRs1("APPO_DESC")
				
		TAX_DOC_CD		= oRs1("TAX_DOC_CD")	' -- 데이타에 따라 서식코드가 달라지므로 

		Call SubCloseRs(oRs1)	

		If EX_RECON_FLG = "Y" Then
			' -- 제1호서식을 읽어온다.
			Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

			' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
			gCursorLocation = adUseClient 

			If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
				Call SaveHTFError(lgsPGM_ID, lgsPGM_NM, TYPE_DATA_NOT_FOUND)
				Exit Function
			End If
		End If

		PrintLog "LoadData Success "
		
		LoadData = True
	End Function				

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pWhereSQL)
		lgoRs1.MoveFirst
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
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm)
		Else
			GetData		= ""
		End If
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
	End Sub


	' ------------------ 조회 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
				lgStrSQL = lgStrSQL & " (	SELECT  "
	            lgStrSQL = lgStrSQL & "		TOP 1 TAX_DOC_CD " & vbCrLf
	            lgStrSQL = lgStrSQL & "		FROM TB_TAX_DOC " & vbCrLf	' 사용자권한별 메뉴뷰 
				lgStrSQL = lgStrSQL & "		WHERE TAX_DOC_CD IN ('A100', 'A138') " & vbCrLf            
	            lgStrSQL = lgStrSQL & " ) TAX_DOC_CD " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY	A  WITH (NOLOCK) " & vbCrLf	' 사용자권한별 메뉴뷰 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

	      Case "H2"
				lgStrSQL = " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_AGENT_INFO	A  WITH (NOLOCK) " & vbCrLf	' 사용자권한별 메뉴뷰 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
		End Select
		PrintLog "SubMakeSQLStatements_WB101MA1 : " & lgStrSQL
	End Sub	
End Class



' ------------------ 메인 함수 --------------------------------
Function MakeHTF_WB101MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, sNowDt, blnError, iSeqNo
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    
    blnError = False
    PrintLog "MakeHTF_WB101MA1 IS RUNNING: "
	lgsPGM_ID	= "WB101MA1"

	Set lgcCompanyInfo = New C_COMPANY_HISTORY		' -- 모든 Include에서 사용할 법인기초 클래스 

	If Not lgcCompanyInfo.LoadData Then Exit Function			' -- 법인기초정보 로드 
		
	'==========================================
	' -- 파일 생성 시작 
	'sNowDt = UNI8Date(Date())
	
	Call InitFileSystem("../../files/" & wgCO_CD ,"HomeTaxFile_" & wgCO_CD &".A100") 	
	

		
	'==========================================
	' -- 법인 신고 공통 생성 및 오류검증 
	sHTFBody = "81"
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	sHTFBody = sHTFBody & "31"

	If Not ChkNotNull(lgcCompanyInfo.FISC_END_DT, "당기종료일자") Then blnError = True
	sHTFBody = sHTFBody & UNI6Date(lgcCompanyInfo.FISC_END_DT)

	If Not ChkBoundary("81,82,84,86,87,88", GetRgstNo42(lgcCompanyInfo.OWN_RGST_NO), "사업자등록번호(4:2)") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO), 13)
	
	If lgsREP_TYPE = "2" Then	' -- 신고구분 
		sHTFBody = sHTFBody & "3"
	Else
		sHTFBody = sHTFBody & "1"
	End If
	sHTFBody = sHTFBody & "0001"	' -- 신고차수 
	sHTFBody = sHTFBody & "0001"	' -- 순차번호 
	sHTFBody = sHTFBody & "8"	' -- 납세자구분 
	
	If Not ChkNotNull(lgcCompanyInfo.HOME_TAX_USR_ID, "사용자ID") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_USR_ID, 20)

	If Not ChkNotNull(lgcCompanyInfo.INCOM_DT, "신고서 제출일") Then blnError = True
	sHTFBody = sHTFBody & UNI6Date(lgcCompanyInfo.INCOM_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.LAW_RGST_NO, "법인등록번호") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.LAW_RGST_NO), 13)
 
	If Not ChkNotNull(lgcCompanyInfo.CO_NM, "법인명") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.CO_NM, "법인명") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.CO_NM, 60)
	
	If Not ChkNotNull(lgcCompanyInfo.REPRE_NM, "대표자명") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.REPRE_NM, "대표자명") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.REPRE_NM, 30)
	
	If Not ChkNotNull(lgcCompanyInfo.CO_ADDR, "사업장소재지") Then blnError = True
	If Not ChkContents(lgcCompanyInfo.CO_ADDR, "사업장소재지") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.CO_ADDR, 70)
	
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_E_MAIL, 30)	' -- 이메일주소 
	
	If Not ChkTelNo(lgcCompanyInfo.TEL_NO, "사업장전화번호") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TEL_NO, 14)	' -- 전화번호 
	
	If Not ChkNotNull(lgcCompanyInfo.IND_CLASS, "업태") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.IND_CLASS, 30)
	
	If Not ChkNotNull(lgcCompanyInfo.IND_TYPE, "업종") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.IND_TYPE, 50)

	If Not ChkNotNull(lgcCompanyInfo.HOME_TAX_MAIN_IND, "주업종코드") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.HOME_TAX_MAIN_IND, 7)
	
	If Not ChkNotNull(lgcCompanyInfo.FISC_START_DT, "당기시작일자") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.FISC_START_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.FISC_END_DT, "당기종료일자") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.FISC_END_DT)

	If Not ChkDateFrTo(lgcCompanyInfo.FISC_START_DT, lgcCompanyInfo.FISC_END_DT, "당기시작일자", "당기종료일자", False) Then blnError = True
	
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_ANY_START_DT) ' -- 수시부과기간 
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_ANY_END_DT)
	
	If Not ChkNotNull(lgcCompanyInfo.HOME_FILE_MAKE_DT, "신고서작성일자") Then blnError = True
	sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.HOME_FILE_MAKE_DT)
	
	If Not ChkContents(lgcCompanyInfo.AGENT_NM, "세무대리인성명") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_NM, 30)	' -- 세무대리인성명 
	
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.RECON_BAN_NO), 6)	' -- 세무대리인관리번호: UNIRemoveDash (2006.03.07수정)
	
	If Not ChkTelNo(lgcCompanyInfo.AGENT_TEL_NO, "세무대리인전화번호") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.AGENT_TEL_NO, 14)	' -- 세무대리인전화번호 
	
	sHTFBody = sHTFBody & UNIChar("1031", 4)	' -- SDS uniERP 세무프로그램코드 
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.AGENT_RGST_NO), 10)	' -- 세무대리인성명 

	sHTFBody = sHTFBody & UNIChar("", 19)	' -- 공란 

	' -- 파일에 기록한다.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	End If


	If lgcCompanyInfo.EX_RECON_FLG = "Y" Then
		Dim blnChk
		
		If Mid(lgcCompanyInfo.RECON_MGT_NO, 4, 1) = "8" Then
			blnChk = True	' -- 검증제외 : 8 이 아닌 경우 그리드중에 조정자관리번호/세무대리인사업자번호가 존재해야 한다.
		Else
			blnChk = False	' -- 검증시작 
		End If
		
		 '-- 2006.03 조정반신청서 생성 
		sHTFBody = "83"
		sHTFBody = sHTFBody & "A218"	' 
		 
		If Not ChkNotNull(lgcCompanyInfo.GetData("W_NAME"), "대표자_성명") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_NAME"), 30)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO1"), "대표자_등록번호") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO1")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_MGT_NO"), "대표자_관리번호") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_MGT_NO")), 6)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO"), "대표자_사업자등록번호") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO2"), "대표자_주민등록번호") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO2")), 13)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_CO_ADDR"), "대표자_사업장소재지") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_CO_ADDR"), 70)

		If Not ChkNotNull(lgcCompanyInfo.GetData("W_HOME_ADDR"), "대표자_주소") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_HOME_ADDR"), 70)

		If Not ChkNotNull(lgcCompanyInfo.REQUEST_DT, "신청일자") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.REQUEST_DT)

		If Not ChkNotNull(lgcCompanyInfo.APPO_NO, "지정번호") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.APPO_NO), 5)

		If Not ChkContents(lgcCompanyInfo.APPO_DESC, "지정조건") Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.APPO_DESC, 50)

		If Not ChkNotNull(lgcCompanyInfo.APPO_DT, "지정일자") Then blnError = True
		sHTFBody = sHTFBody & UNI8Date(lgcCompanyInfo.APPO_DT)
		
		sHTFBody = sHTFBody & UNIChar("", 8) & vbCrLf	' -- 공란 

		If Not blnError Then
			PrintLog "WriteLine2File : " & sHTFBody
			Call PushRememberDoc(sHTFBody)	' -- 바로 출력하지 않고 기억시킨다(inc_HomeTaxFunc.asp에 정의)
		End If

		If blnChk = False Then
			If lgcCompanyInfo.GetData("W_MGT_NO") = lgcCompanyInfo.RECON_MGT_NO Or lgcCompanyInfo.AGENT_RGST_NO = lgcCompanyInfo.GetData("W_RGST_NO") Then
				blnChk = True	' -- 일치 발견 
			End If
		End If
		
		' -- 다음행으로..
		lgcCompanyInfo.MoveNext 
		iSeqNo = 1	: sHTFBody = ""	' -- 초기화 
		
		' -- 조정반 그리드 
		Do Until lgcCompanyInfo.EOF 
			 '-- 2006.03 조정반신청서 생성 
			sHTFBody = sHTFBody & "84"
			sHTFBody = sHTFBody & "A218"	' 
			
			sHTFBody = sHTFBody & UNINumeric(iSeqNo, 6, 0)
			
			If Not ChkNotNull(lgcCompanyInfo.GetData("W_NAME"), "구성원_성명") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_NAME"), 30)
			
			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO1"), "구성원_등록번호") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO1")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_MGT_NO"), "구성원_관리번호") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_MGT_NO")), 6)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO"), "구성원_사업자등록번호") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_RGST_NO2"), "구성원_주민등록번호") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.GetData("W_RGST_NO2")), 13)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_CO_ADDR"), "구성원_사업장소재지") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_CO_ADDR"), 70)

			If Not ChkNotNull(lgcCompanyInfo.GetData("W_HOME_ADDR"), "구성원_주소") Then blnError = True
			sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.GetData("W_HOME_ADDR"), 70)

			sHTFBody = sHTFBody & UNIChar("", 23) & vbCrLf	' -- 공란 

			' -- 검증필요 
			If blnChk = False Then
				If lgcCompanyInfo.GetData("W_MGT_NO") = lgcCompanyInfo.RECON_MGT_NO Or lgcCompanyInfo.AGENT_RGST_NO = lgcCompanyInfo.GetData("W_RGST_NO") Then
					blnChk = True	' -- 일치 발견 
				End If
			End If
			
			iSeqNo = iSeqNo + 1
			lgcCompanyInfo.MoveNext 
		Loop

		If blnChk = False Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, "", "세무대리인 기본정보의 세무대리인사업자번호의 4번째 자리가 '8'이 아닐 경우 조정반에 대한 사항(그리드)중에 관리번호/사업자 번호에 일치하는 데이타가 존재해야 합니다.")
		End If
		
		If Not blnError Then
			PrintLog "WriteLine2File : " & sHTFBody
			Call PushRememberDoc(sHTFBody)	' -- 바로 출력하지 않고 기억시킨다(inc_HomeTaxFunc.asp에 정의)
		End If

		' -- 레코드셋 닫기 
		Call lgcCompanyInfo.CloseRs
	End If
	
	' -- 법인공통은 뒤에서도 쓰므로 메모리해제를 안함 
End Function


%>
