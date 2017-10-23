<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제54호 주식등 변동상황명세서(갑)
'*  3. Program ID           : W9109MA1
'*  4. Program Name         : W9109MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_54

Set lgcTB_54 = Nothing ' -- 초기화 

Class C_TB_54
	' -- 테이블의 컬럼변수 
	Dim TAX_OFFICE
	Dim INCOM_DT
	Dim W5
	Dim W6
	
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

	'----------- 멀티 행 지원 ------------------------
	Function Find(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 2
				lgoRs2.Find pWhereSQL
			Case 3
				lgoRs3.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
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
			Case 2
			    If Not lgoRs1.EOF Then
				   lgoRs2.MoveFirst
				End if   
			Case 3
				lgoRs3.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 2
			    If Not lgoRs1.EOF Then
				   lgoRs2.MoveNext
				End if   
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
	
	Function RecordCount(Byval pType)
		Select Case pType
			Case 3
				RecordCount = lgoRs3.RecordCount
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
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
	            lgStrSQL = lgStrSQL & " B.TAX_OFFICE, B.INCOM_DT, B.SUBMIT_FLG  " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.RECON_MGT_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE ''" & vbCrLf
	            lgStrSQL = lgStrSQL & "   END RECON_MGT_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.OWN_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "   END AGENT_RGST_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_NM" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.CO_NM " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END AGENT_NM "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN C.W_NAME" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.REPRE_NM " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END REPRE_NM "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN C.W_CO_ADDR" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.CO_ADDR " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END CO_ADDR "& vbCrLf
	            lgStrSQL = lgStrSQL & " , CASE B.SUBMIT_FLG WHEN '1' THEN B.AGENT_TEL_NO" & vbCrLf
	            lgStrSQL = lgStrSQL & "		ELSE B.TEL_NO " & vbCrLf
	            lgStrSQL = lgStrSQL & "	  END TEL_NO "& vbCrLf
	            lgStrSQL = lgStrSQL & " , B.OWN_RGST_NO, B.CO_NM, B.REPRE_NM ,HOME_TAX_USR_ID" & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54H	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & "		INNER JOIN TB_COMPANY_HISTORY	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN TB_AGENT_INFO C  WITH (NOLOCK) ON B.CO_CD=C.CO_CD AND B.FISC_YEAR=C.FISC_YEAR AND B.REP_TYPE=C.REP_TYPE AND C.W_TYPE='대표이사' " & vbCrLf	' 서식3호 

				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54D1	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54D2	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
							
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9109MA1
	Dim A101

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, arrTmp1(1), arrTmp2(1)
    Dim iSeqNo, iDx, iDtlCnt,dblNum1,dblAmtRate1,dblNum2
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9109MA1"

	Set lgcTB_54 = New C_TB_54		' -- 해당서식 클래스 
	
	If Not lgcTB_54.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9109MA1

	'==========================================
	' -- 제54호 주식등 변동상황명세서(갑) 오류검증 
	' -- 1. 기본사항(83A131)
	sHTFBody = "A"
	sHTFBody = sHTFBody & UNIChar("40", 2)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	
	
	  ' -- 제8호 갑 공제감면세액명세서 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A100  = new C_TB_1	' -- W8101MA1_HTF.asp 에 정의됨 
							
		
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제1호 법인세과세표준및세액신고서(A100)", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
            '- 일반법인의 비영리 법인(법인구분 : 50, 60, 70)인 경우 입력되면 오류- 법인세과세표준및세액신고서(A100)의 주식변동자료매체로제출여부가 'Y' 인 경우 입력되면 오류 
		     If cDataExists.A100.W2 = "50" Or cDataExists.A100.W2 = "60" Or cDataExists.A100.W2="70" Then
		        blnError = True	
		        Call SaveHTFError(lgsPGM_ID, cDataExists.A100.W2, UNIGetMesg("일반법인의 비영리 법인(법인구분 : 50, 60, 70)인 경우 입력되면 안됩니다", "",""))
		        Exit Function
		     End IF
		    
		End If	
						
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A100 = Nothing
		Set cDataExists = Nothing	' -- 메모리해제 
	
		 
		 If lgcCompanyInfo.EX_54_FLG = "Y"  Then   ' 주식매체제출 
		   blnError = True	
		   Call SaveHTFError(lgsPGM_ID, lgcTB_3.W02, UNIGetMesg("법인세과세표준및세액신고서(A100)의 주식변동자료매체로제출여부가 'Y' 인 경우 입력되면 안됩니다", "",""))
		   Exit Function
		End IF
		
		
	
	If  ChkNotNull(lgcTB_54.GetData(1, "TAX_OFFICE"), "세무서코드") Then 
	
	    If Not ChkNumeric(lgcTB_54.GetData(1, "TAX_OFFICE"),"세무서코드") Then
	      blnError = True	
	    End If  
	   
	Else
	    blnError = True	
	End If   
	 sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)
		
	If  ChkNotNull(lgcTB_54.GetData(1, "INCOM_DT"), "제출년월일") Then
	    If Not ChkDate(lgcTB_54.GetData(1, "INCOM_DT"),"제출년월일") Then
	        blnError = True	
	    End If
	Else
	     blnError = True	    
	End if    
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "INCOM_DT"))
	
		
	If  ChkNotNull(lgcTB_54.GetData(1, "SUBMIT_FLG"), "제출자구분") Then
	    '- ‘1’ : 세무대리인, ‘2’ : 납세자 
	    If Not ChkBoundary("1,2" , lgcTB_54.GetData(1, "SUBMIT_FLG") , "제출자구분") Then
	       blnError = True	
	    End If
	Else
	    blnError = True	
	End if
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "SUBMIT_FLG"), 1)        
	
	
	''- 납세자가 제출하는 경우에는 공란- 입력된 경우 법인세신고공통(A100)의 세무대리인관리번호와 일치 
	If lgcTB_54.GetData(1, "SUBMIT_FLG") = "1" Then
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "RECON_MGT_NO"), 6)
	Else
		sHTFBody = sHTFBody & UNIChar("", 6)
	End If
		
	 '- 제출자_사업자등록번호 :숫자 Check- 제출자 사업자등록번호와 납세자 사업자등록번호 일치 여부 Check- 
			    '세무대리인관리번호가 입력된 경우 :  법인세신고공통(A100)의 세무대리인 사업자등록번호와 일치 
			    '세무대리인관리번호가 공란인 경우 :  법인세신고공통(A100)의 납세자ID와 일치 
	if lgcTB_54.GetData(1, "RECON_MGT_NO") <>"" Then
			If  ChkNotNull(lgcTB_54.GetData(1, "AGENT_RGST_NO"), "제출자_사업자등록번호") Then 
	
			      If Not  ChkNumeric(UNIRemoveDash(lgcTB_54.GetData(1, "AGENT_RGST_NO")), "제출자_사업자등록번호" ) Then  blnError = True	
			   		   
			Else
			    blnError = True	
			End if    
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "AGENT_RGST_NO")), 10)
	Else
	        sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10)
	End if		
	
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "AGENT_NM"), "제출자_상호") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "AGENT_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "REPRE_NM"), "제출자_성명") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "REPRE_NM"), 30)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "CO_ADDR"), "제출자_주소") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "CO_ADDR"), 80)
		
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TEL_NO"), 15)
		
	sHTFBody = sHTFBody & UNIChar("101", 3)	' 사용한한글코드 
	
	sHTFBody = sHTFBody & UNIChar("", 591)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False
	sHTFBody = ""
	
	
	' -- 1. 자본금변동상황 
	sHTFBody = "B"
	sHTFBody = sHTFBody & UNIChar("40", 2)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	
	If  ChkNotNull(lgcTB_54.GetData(1, "TAX_OFFICE"), "세무서코드") Then 
	
	    If Not ChkNumeric(lgcTB_54.GetData(1, "TAX_OFFICE"),"세무서코드") Then
	      blnError = True	
	    End If  
	   
	Else
	    blnError = True	
	End If   
	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)	' -- 세무서코드 
	
	
	

	If  ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "신고법인_사업자등록번호") Then 
	
	      If Not  ChkNumeric(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), "신고법인_사업자등록번호" ) Then  blnError = True	
			   		   
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10)
	
	
	If Not ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "신고법인_법인명") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "CO_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54.GetData(1, "OWN_RGST_NO"), "신고법인_대표자성명") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "REPRE_NM"), 30)

	' -- 2005.03.24 개정서식 
'	상장(등록)변경일    
'    - 공백,’00000000’허용, 
'    - 아닐 경우 날짜형식 Check, 상장(등록)변경일 >= 두번째사업년도(시작일자), 
'                                상장(등록)변경일 <= 두번째사업년도(종료일자)
'    합병. 분할일    
'    - 공백,’00000000’허용, 
'    - 아닐 경우 날짜형식 Check, 합병. 분할일 >= 두번째사업년도(시작일자), 
'                                합병. 분할일 <= 두번째사업년도(종료일자)
		
	If lgcTB_54.GetData(1, "W4") <> "" Then 
		If lgcTB_54.GetData(1, "W4") <> "00000000" Then
			If Not ChkDate(lgcTB_54.GetData(1, "W4"),"상장(등록)변경일") Then
				blnError = True	
			End If

			If 	CDate(lgcTB_54.GetData(1, "W4")) = CDate(lgcTB_54.GetData(1, "W6_2"))-1 Then
			ElseIf 	CDate(lgcTB_54.GetData(1, "W4")) = CDate(lgcTB_54.GetData(1, "W6_1")) Then
			Else
				Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W4") & " , " & lgcTB_54.GetData(1, "W6_2"), "상장(등록)변경일 >= 두번째사업연도(시작일자) 또는 상장(등록)변경일 <= 두번째사업연도(종료일자)") 
				blnError = True
			End If
			
		End If
	    sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W4"))
	Else
	      sHTFBody = sHTFBody & UNIChar("", 8)    
	End If
	
		

	If lgcTB_54.GetData(1, "W5") <> "" Then 
		If lgcTB_54.GetData(1, "W5") <> "00000000" Then
			If Not ChkDate(lgcTB_54.GetData(1, "W5"),"합병.분할일") Then
			    blnError = True	
			End If
		End If

		If 	CDate(lgcTB_54.GetData(1, "W5")) = CDate(lgcTB_54.GetData(1, "W6_2"))-1 Then
		ElseIf 	CDate(lgcTB_54.GetData(1, "W5")) = CDate(lgcTB_54.GetData(1, "W6_1")) Then
		Else
			Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W5") & " , " & lgcTB_54.GetData(1, "W6_2"), "합병.분할일 >= 두번째사업연도(시작일자) 또는 합병.분할일 <= 두번째사업연도(종료일자)") 
			blnError = True
		End If

	    sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W5"))
	Else
	    sHTFBody = sHTFBody & UNIChar("", 8)
	End If


' -- 2006.03.24 개정으로 제거	
'	If  ChkNotNull(lgcTB_54.GetData(1, "W6_1"), "사업연도(시작일자)") Then
'	    If  UNI8Date(lgcCompanyInfo.FISC_START_DT) <>UNI8Date(lgcTB_54.GetData(1, "W6_1")) Then
'	        Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W6_1"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "사업연도(시작일자)", "기초정보의사업연도(시작일자)")) 
'			blnError = True	
'	    End If 
'	Else
'	    blnError = True	
'	End if
	
	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W6_1"))

' -- 2006.03.24 개정으로 제거	
'	If  ChkNotNull(lgcTB_54.GetData(1, "W6_2"), "사업연도(종료일자)") Then
'	    If  UNI8Date(lgcCompanyInfo.FISC_END_DT) <>UNI8Date(lgcTB_54.GetData(1, "W6_2")) Then
'	        Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(1, "W6_2"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "사업연도(종료일자)", "기초정보의사업연도(종료일자)"))
'			blnError = True	
'	    End If 
'	Else
'	    blnError = True	
'	End if
	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(1, "W6_2"))
		
	' -- 그리드 저장 
	
	If  ChkNotNull(lgcTB_54.GetData(2, "W10"), "당기초_총발행주식수_보통주") Then 
	    '당기초_총발행주식수_보통주+ 당기초_총발행주식수_우선주	- 항목(21)기초주식수_합계와 일치 
	    dblNum1 = Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	Else
	    blnError = True	
	End if    
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	arrTmp1(0) = lgcTB_54.GetData(2, "W11")
	arrTmp1(1) = lgcTB_54.GetData(2, "W13")
	
	Call lgcTB_54.MoveNext(2)
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "당기초_총발행주식수_우선주") Then blnError = True	
	  dblNum1 = dblNum1 + Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	Call lgcTB_54.Find(2, "SEQ_NO=13")	' 기말로 감 
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "당기말_총발행주식수_보통주") Then blnError = True	
	dblNum2 = Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)

	arrTmp2(0) = lgcTB_54.GetData(2, "W11")
	arrTmp2(1) = lgcTB_54.GetData(2, "W13")
	
	Call lgcTB_54.MoveNext(2)
	
	If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "당기말_총발행주식수_우선주") Then blnError = True	
	dblNum2= dblNum2 + Unicdbl(lgcTB_54.GetData(2, "W10"),0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	If Not ChkNotNull(arrTmp1(0), "당기초_주당액면가액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp1(0), 8, 0)
	
	If Not ChkNotNull(arrTmp2(0), "당기말_주당액면가액") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp2(0), 8, 0)
	
	If Not ChkNotNull(arrTmp1(1), "당기초_자본금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp1(1), 16, 0)
	
	If Not ChkNotNull(arrTmp2(1), "당기말_자본금") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(arrTmp2(1), 16, 0)
	
	Call lgcTB_54.MoveFist(2)
	Call lgcTB_54.Find(2, "SEQ_NO=3")	' 기초 다음으로 감 
	
	'PrintLog "SEQ_NO=3=" & lgcTB_54.EOF(2)
	'Response.End 
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W7"), "변동_일자1") Then blnError = True	
	If lgcTB_54.GetData(2, "W7") = "" Then 
	   sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(2, "W7"))
	Else
	   sHTFBody = sHTFBody & UNIChar("", 8)
	End If   
	      
	
	'If Not ChkBoundary("01,02,03,04,05,06,07,08,09", lgcTB_54.GetData(2, "W8"), "변동_원인코드1") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(2, "W8"), 2)
	if lgcTB_54.GetData(2, "W7") <> "" Then
	   '- 날짜형식 Check- 일자 >= 사업년도(시작일자), 일자 <= 사업년도(종료일자)
	   ' 일자가 공백(SPACE)인 경우 원인코드, 종류는 공백이어야 하고, 주식수, 주당액면가액, 주당발행(인수)가액, 증가(감소)자본금은 0
	    sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W9"), 1, 0)   '변동_종류1
	Else
	    sHTFBody = sHTFBody & UNIChar("", 1)
	End If     
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W10"), "변동_주식수1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W11"), "변동_주당액면가액1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W11"), 8, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W13"), "변동_증가(감소)자본금1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W12"), 8, 0)
	
	'If Not ChkNotNull(lgcTB_54.GetData(2, "W13"), "변동_증가(감소)자본금1") Then blnError = True	
	sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W13"), 16, 0)
		 
	
	Call lgcTB_54.MoveNext(2)

	For iDx = 2 To 10
		If lgcTB_54.GetData(2, "W7") = "" Then 
	       sHTFBody = sHTFBody & UNI8Date(lgcTB_54.GetData(2, "W7"))
	    Else
	       sHTFBody = sHTFBody & UNIChar("", 8)
	   End If   
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(2, "W8"), 2)
		if lgcTB_54.GetData(2, "W7") <> "" Then
	   '- 날짜형식 Check- 일자 >= 사업년도(시작일자), 일자 <= 사업년도(종료일자)
	   ' 일자가 공백(SPACE)인 경우 원인코드, 종류는 공백이어야 하고, 주식수, 주당액면가액, 주당발행(인수)가액, 증가(감소)자본금은 0
	    sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W9"), 1, 0)   '변동_종류1
		Else
	    sHTFBody = sHTFBody & UNIChar("", 1)
		End If  
	
		   
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W10"), 13, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W11"), 8, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W12"), 8, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(2, "W13"), 16, 0)
		
		Call lgcTB_54.MoveNext(2)
	Next
	
	' 주식수변동상황건수 
	iDtlCnt = lgcTB_54.RecordCount(3)
	sHTFBody = sHTFBody & UNINumeric(iDtlCnt, 6, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 6)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
	
	' -- 주식수변동상황 
	iSeqNo = 1	
	
	Do Until lgcTB_54.EOF(3) 

		sHTFBody = sHTFBody & "C"
		sHTFBody = sHTFBody & UNIChar("40", 2)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(1, "TAX_OFFICE"), 3)	' -- 세무서코드 
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(1, "OWN_RGST_NO")), 10) ' -- 사업자번호 
        If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and  UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
		     If Not ChkBoundary("1,2,3", lgcTB_54.GetData(3, "W17_1"), "내용구분") Then blnError = True
		End if     
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W17_1"), 1)
	
	    If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
	    	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W17"), 1)
	    
	    Else
	        sHTFBody = sHTFBody & UNIChar("", 1)
	    End If	
		
	
	    If   UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) <> 1  and  UNICDbl(lgcTB_54.GetData(3, "W17_1"),0)  <> 2 Then 
	    	sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W18"), 40)
	    Else
	        sHTFBody = sHTFBody & UNIChar(" ", 40)
	    End If	
		if  UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then 
		    sHTFBody = sHTFBody & UNIChar("AAAAAAAAAAAAA", 13) 
		Elseif  UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 2 Then     
		    sHTFBody = sHTFBody & UNIChar("BBBBBBBBBBBBB", 13) 
		Else
		
			If  Not ChkNotNull(lgcTB_54.GetData(3, "W19"), "주주_주민(사업자)등록번호") Then
			   blnError = True	
			    Stmp= "주주_주민(사업자)등록번호 는 " 
				Stmp = sTmp  &  " 우리사주소계 :	 		'CCCCCCCCCCCCC' "
				Stmp = sTmp  &  " 주주구분이 4, 5, 6이고 해당번호를 알 수 없을 때는  " 
				Stmp = sTmp  &  "  1) 개인단체(주주구분：4)는 	'DDDDDDDDDDDDD' " 
				Stmp = sTmp  &  "    ※ 해당번호를 알 수 없는 개인단체가 여러건 존재하는 경우  " 
				Stmp = sTmp  &  "       'DDDDDDDDDD+일련번호(3자리)'로 구분하여 입력  " 
				Stmp = sTmp  &  "       (예)DDDDDDDDDD001, DDDDDDDDDD002,…  "
				Stmp = sTmp  &  "  2) 외국투자자(주주구분：5)는 	'EEEEEEEEEEEEE'  "
				Stmp = sTmp  &  "    ※ 해당번호를 알 수 없는 외국투자자가 여러건 존재하는 경우  " 
				Stmp = sTmp  &  "       위와 같은 방법으로 구분하여 입력 (예)EEEEEEEEEE001,…   "
				Stmp = sTmp  &  " 3) 외국법인(주주구분：6)은 	'FFFFFFFFFFFFF' " 
				Stmp = sTmp  &  "    ※ 해당번호를 알 수 없는 외국법인이 여러건 존재하는 경우  " 
				Stmp = sTmp  &  "       위와 같은 방법으로 구분하여 입력 (예)FFFFFFFFFF001,…   " 
				      
				Stmp = sTmp  &  "     그외에는 정상입력하고   위와 같이 입력해 주십시오"
			   
			   
			   Call SaveHTFError(lgsPGM_ID, lgcTB_54.GetData(3, "W19"), UNIGetMesg(Stmp, "",""))
			   
			End If    
			sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54.GetData(3, "W19")), 13)
		End if	
	
	
	
	    '주식등변동상황명세서_자본금변동상황(84A131)의 당기초_총발행주식수와 일치 
		If  ChkNotNull(lgcTB_54.GetData(3, "W20"), "기초주식수") Then 
		    If dblNum1  <> UNICDbl(lgcTB_54.GetData(3, "W20"),0)  AND    UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then  '내용구분이 1이면 
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, dblNum1, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기초주식수", "당기초_총발행주식수"))
		    End IF   
		Else
			blnError = True	
		End IF	
		
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W20"), 13, 0)
		'지분율 = 기초주식수 / 주식등변동상황명세서_자본금변동상황(84A131)의 당기초_총발행주식수 x 100
		If  ChkNotNull(lgcTB_54.GetData(3, "W21"), "지분율") Then 
		    If UNICDBl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then '합계 
		   
					 If dblNum1 = 0 Then 
					     dblAmtRate1 = 0
					 Else
					      dblAmtRate1 = (Unicdbl(lgcTB_54.GetData(3, "W20"),0) /  dblNum1 )* 100

					 End if  
					If Unicdbl(lgcTB_54.GetData(3, "W21"),0) <> dblAmtRate1  Then 
					   blnError = True	
					   Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W21"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "지분율", "기초주식수" & Unicdbl(lgcTB_54.GetData(3, "W20"),0) &" / 주식등변동상황명세서_자본금변동상황(84A131)의 당기초_총발행주식수" & dblNum1 &"x 100"))
                	End If 
		      End If
		   
		Else
			 blnError = True	
		End IF	 
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W21"), 5, 2)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W22"), "증가_양수") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W22"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W23"), "증가_유상증자") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W23"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W24"), "증가_무상증자") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W24"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W25"), "증가_상속") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W25"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W26"), "증가_증여") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W26"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W27"), "증가_출자전환") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W27"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W28"), "증가_기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W28"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W29"), "감소_양도") Then blnError = True	
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W29"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W30"), "감소_상속") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W30"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W31"), "감소_증여") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W31"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W32"), "감소_감자") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W32"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W33"), "감소_기타") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W33"), 13, 0)
		
		'- 내용구분이 '1'(합계)· 당기말 총발행주식수와 일치· 
		'내용구분 '2'(소액주주소계) + 내용구분 '3'(개별주주)와 일치 · 
		'항목(16)기말의 항목(11)주식수와 일치 
		' 기말주식수 = 기초주식수 + 증가_양수 + 증가_유상증자 + 증가_무상증자 + 증가_상속 + 증가_증여 + 증가_출자전환 + 증가_기타 - 감소_양도 - 감소_상속 - 감소_증여 - 감소_감자 - 감소_기타 
		If  ChkNotNull(lgcTB_54.GetData(3, "W34"), "기말주식수") Then
		   ' dblNum2  -  당기말 총발행주식수 
		     If UNICDbl(lgcTB_54.GetData(3, "W34"),0) <> dblNum2  AND    UNICDbl(lgcTB_54.GetData(3, "W17_1"),0) = 1 Then  '내용구분이 1이면 
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W34"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기말주식수", "당기말_총발행주식수"))
		       
		     End IF
		     
		     If Unicdbl(lgcTB_54.GetData(3, "W34"),0) <> UNICDbl(lgcTB_54.GetData(3, "W20"),0) + UNICDbl(lgcTB_54.GetData(3, "W22"),0)+ UNICDbl(lgcTB_54.GetData(3, "W23"),0) + UNICDbl(lgcTB_54.GetData(3, "W24"),0)  + UNICDbl(lgcTB_54.GetData(3, "W25"),0) _
		        + UNICDbl(lgcTB_54.GetData(3, "W26"),0) + UNICDbl(lgcTB_54.GetData(3, "W27"),0) + UNICDbl(lgcTB_54.GetData(3, "W28"),0) - UNICDbl(lgcTB_54.GetData(3, "W29"),0)  - UNICDbl(lgcTB_54.GetData(3, "W30"),0) _
		        - UNICDbl(lgcTB_54.GetData(3, "W30"),0) - UNICDbl(lgcTB_54.GetData(3, "W31"),0) - UNICDbl(lgcTB_54.GetData(3, "W32"),0) - UNICDbl(lgcTB_54.GetData(3, "W33"),0) Then
		          
		         blnError = True	
			     Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W34"),0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "기말주식수", " 기초주식수 + 증가_양수 + 증가_유상증자 + 증가_무상증자 + 증가_상속 + 증가_증여 + 증가_출자전환 + 증가_기타 - 감소_양도 - 감소_상속 - 감소_증여 - 감소_감자 - 감소_기타"))
		     End If
		Else     
		    blnError = True	
		End If    
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W34"), 13, 0)
		
		If Not ChkNotNull(lgcTB_54.GetData(3, "W35"), "기말지분율") Then blnError = True	
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54.GetData(3, "W35"), 5, 2)
		
		If lgcTB_54.GetData(3, "W36") <> "" Then
		   If Not ChkBoundary("00,01,02,03,04,05,06,07,08,09", lgcTB_54.GetData(3, "W36"), "지배주주와의관계") Then blnError = True
		   If (lgcTB_54.GetData(3, "W17") = "2" Or lgcTB_54.GetData(3, "W17") = "3" Or lgcTB_54.GetData(3, "W17") = "4" Or lgcTB_54.GetData(3, "W17") = "6" ) And lgcTB_54.GetData(3, "W36") <> "09" Then 
		      '주주구분이 ‘2’,’3’,’4’,’6’ 이면 지배주주와의 관계는 ‘09’가 아니면 오류 
		        blnError = True	
			    Call SaveHTFError(lgsPGM_ID, Unicdbl(lgcTB_54.GetData(3, "W36"),0), UNIGetMesg("주주구분이 ‘2’,’3’,’4’,’6’ 이면 지배주주와의 관계는 ‘09’가 아니면 오류입니다", "", ""))
		   End IF
		End If   
		sHTFBody = sHTFBody & UNIChar(lgcTB_54.GetData(3, "W36"), 2)
	
			
		sHTFBody = sHTFBody & UNIChar("", 519) & vbCrLf

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_54.MoveNext(3)	' -- 1번 레코드셋 
	Loop



	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_54 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9109MA1 : " & lgStrSQL
End Sub
%>
