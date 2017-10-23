<%
'======================================================================================================
'*  1. Function Name        : 전산1호 전산운용조직명세서 
'*  3. Program ID           : W9123MA1
'*  4. Program Name         : W9123MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_JS1

Set lgcTB_JS1 = Nothing	' -- 초기화 

Class C_TB_JS1
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W2_ETC
	Dim W3
	Dim W3_ETC
	Dim W4
	Dim W4_ETC
	Dim W4_1
	Dim W5
	Dim W5_ETC
	Dim W6
	Dim W6_ETC
	Dim W7_1
	Dim W7_2
	Dim W8
	Dim W9_1
	Dim W9_2
	Dim W9_3
	Dim W9_4
	Dim W9_5
	Dim W9_6
	Dim W9_6__ETC
	Dim W10_1
	Dim W10_2
	Dim W10_3
	Dim W10_4
	Dim W10_5
	Dim W10_6
	Dim W10_7
	Dim W10_8
	Dim W10_9
	Dim W10_9__ETC
		
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

		W1			= oRs1("W1")
		W2			= oRs1("W2")	
		W2_ETC		= oRs1("W2_ETC")
		W3			= oRs1("W3")
		W3_ETC		= oRs1("W3_ETC")
		W4			= oRs1("W4")
		W4_ETC		= oRs1("W4_ETC")
		W4_1		= oRs1("W4_1")
		W5			= oRs1("W5")
		W5_ETC		= oRs1("W5_ETC")
		W6			= oRs1("W6")
		W6_ETC		= oRs1("W6_ETC")
		W7_1		= oRs1("W7_1")
		W7_2		= oRs1("W7_2")
		W8			= oRs1("W8")
		W9_1		= oRs1("W9_1")
		W9_2		= oRs1("W9_2")
		W9_3		= oRs1("W9_3")
		W9_4		= oRs1("W9_4")
		W9_5		= oRs1("W9_5")
		W9_6		= oRs1("W9_6")
		W9_6__ETC	= oRs1("W9_6__ETC")
		W10_1		= oRs1("W10_1")
		W10_2		= oRs1("W10_2")
		W10_3		= oRs1("W10_3")
		W10_4		= oRs1("W10_4")
		W10_5		= oRs1("W10_5")
		W10_6		= oRs1("W10_6")
		W10_7		= oRs1("W10_7")
		W10_8		= oRs1("W10_8")
		W10_9		= oRs1("W10_9")
		W10_9__ETC	= oRs1("W10_9__ETC")
		
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
				lgStrSQL = lgStrSQL & " FROM TB_JS1	A  WITH (NOLOCK) " & vbCrLf	' 서식1호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
		End Select
		
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
  	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9123MA1
	Dim A100

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9123MA1()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx, iMaxRows
    Dim iLoopMax, sHTFBody, sNowDt, blnError, oRs2, sTmp, cDataExists
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9123MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9123MA1"
	
	Set lgcTB_JS1 = New C_TB_JS1		' -- 해당서식 클래스 
	
	If Not lgcTB_JS1.LoadData	Then Exit Function		' -- 제1호 서식 로드 

	'==========================================
	
	
	' -- 전산1호 전산운용조직명세서 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)	' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkBoundary("1,2,3,4", lgcTB_JS1.W1, "회계프로그램(시스템)사용현황: " & lgcTB_JS1.W1 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W1, 1)
	' OS(운영체제)
	
	If blnError Then Response.Write "회계프로그램(시스템)사용현=" & blnError & vbCrLf
	
	
	If lgcTB_JS1.W2 <> "" Then
		If Not ChkBoundary("1,2,3,4,5,6", lgcTB_JS1.W2, "OS(운영체제)" & lgcTB_JS1.W2 & " " ) Then blnError = True
		sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W2, 1)
		if UNIChar(lgcTB_JS1.W2, 1) = "6"  and Trim(lgcTB_JS1.W2_ETC) ="" then 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W2, UNIGetMesg(TYPE_CHK_NULL, "OS(운영체제)가 6-기타이면 OS 기타명 ",""))
		   
		     blnError = True
		End if     
	End If	
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W2_ETC, 50)
	'프로그램 언어 
	If lgcTB_JS1.W3 <> "" Then
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_JS1.W3, "프로그램 언어: " & lgcTB_JS1.W3 & " " ) Then blnError = True
		
	
		'프로그램 언어 (기타)
		if UNIChar(lgcTB_JS1.W3, 1) = "8"  and Trim(lgcTB_JS1.W3_ETC) ="" then 
		     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W3, UNIGetMesg(TYPE_CHK_NULL, "프로그램 언어 선택이 (6)기타이면 OS 프로그램 언어 (기타) ",""))
		     blnError = True
		End if   
	End If
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W3, 1)	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W3_ETC, 50)
	'DBMS 
	
	If lgcTB_JS1.W4 <> "" Then 
		If Not ChkBoundary("1,2,3,4,5,6,7,8", lgcTB_JS1.W4, "DBMS: " & lgcTB_JS1.W4 & " " ) Then blnError = True
	END IF	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4, 1)
	
	'DBMS (기타)
	if UNIChar(lgcTB_JS1.W4, 1) = "8"  and Trim(lgcTB_JS1.W4_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W4, UNIGetMesg(TYPE_CHK_NULL, "DBMS  선택이 (8)기타이면 DBMS (기타) ",""))
	     blnError = True
	End if   
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4_ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W4_1, 30)

	if UNIChar(lgcTB_JS1.W1, 1) = "3"  then 
	
	   If  Trim(lgcTB_JS1.W5) =""  then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W5, UNIGetMesg(TYPE_CHK_NULL, "회계프로그램(시스템)사용현황  선택이 (3)ERP 이면 ERP",""))
	     blnError = True
     	Else
		   If Not ChkBoundary("1,2,3,4,5", lgcTB_JS1.W5, "ERP: " & lgcTB_JS1.W5 & " " ) Then blnError = True
'		   blnError = True '  2006.03 수정 IF 둘다 False 이다.
		End if   
	End if 
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W5, 1)
	if UNIChar(lgcTB_JS1.W5, 1) = "5"  and Trim(lgcTB_JS1.W5_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W5, UNIGetMesg(TYPE_CHK_NULL, "ERP  선택이 (5)기타이면 ERP(기타) ",""))
	     blnError = True
	End if   
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W5_ETC, 50)
	
	if UNIChar(lgcTB_JS1.W1, 1) = "4" then 
	   if    Trim(lgcTB_JS1.W6) ="" then
	         Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W6, UNIGetMesg(TYPE_CHK_NULL, "회계프로그램(시스템)사용현황  선택이 (4)상업용 회계프로그램이면 상업용 회계프로그램",""))
	         blnError = True
	         	
	   Else
	          If Not ChkBoundary("1,2,3,4,5", lgcTB_JS1.W6, "상업용 회계프로그램: " & lgcTB_JS1.W6 & " " ) Then blnError = True   
	           
	            	
	   end if      

	End if 
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W6, 1)
	
	if UNIChar(lgcTB_JS1.W6, 1) = "5"  and Trim(lgcTB_JS1.W6_ETC) ="" then 
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W6, UNIGetMesg(TYPE_CHK_NULL, "상업용 회계프로그램  선택이 (5)기타이면 상업용 회계프로그램(기타) ",""))
	     blnError = True
	End if 
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W6_ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W7_1, 50)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W7_2, 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	sHTFBody = sHTFBody & UNIChar("", 50)
	sHTFBody = sHTFBody & UNIChar("", 50)
	
	
	If Not ChkBoundary("1,2", lgcTB_JS1.W8, "전자상거래 유무: " & lgcTB_JS1.W8 & " " ) Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W8, 1)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_1, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_2, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_3, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_4, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_5, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_6, 1)
	
	if UNIChar(lgcTB_JS1.W9_6, 1) = "Y" and Trim(lgcTB_JS1.W9_6__ETC) = "" then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W9_6__ETC, UNIGetMesg(TYPE_CHK_NULL, "전자상거래유형이 기타이면 전자상거래기타명 ",""))
	     blnError = True
	End if     
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W9_6__ETC, 50)
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_1, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_2, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_3, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_4, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_5, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_6, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_7, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_8, 1)
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_9, 1)
	if UNIChar(lgcTB_JS1.W10_9, 1) = "Y" and Trim(lgcTB_JS1.W10_9__ETC) = "" then
	     Call SaveHTFError(lgsPGM_ID, lgcTB_JS1.W10_9__ETC, UNIGetMesg(TYPE_CHK_NULL, "단위업무시스템종류가 기타이면 단위업무시스템기타명 ",""))
	     blnError = True
	End if  
	
	sHTFBody = sHTFBody & UNIChar(lgcTB_JS1.W10_9__ETC, 50)

	sHTFBody = sHTFBody & UNIChar("", 42)	' -- 공란 
	
	' ----------- 
		
	' -- 파일에 기록한다.
	If Not blnError Then
		PrintLog "WriteLine2File : " & sHTFBody
		Call WriteLine2File(sHTFBody)
	Else
		Response.End 
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	'Set lgcTB_JS1 = Nothing	' -- 메모리해제  <-- W8101MA1_HTF에서 사용함 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9123MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
  
	  Case "O" '-- 외부 참조 금액 
			
			lgStrSQL = ""
			' -- 검증을 위해 데이타 존재 체크 
			
	End Select
	PrintLog "SubMakeSQLStatements_W9123MA1 : " & lgStrSQL
End Sub

%>
