<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제20호 감가상각비명세서 합계표 
'*  3. Program ID           : W3129MA1
'*  4. Program Name         : W3129MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_20

Set lgcTB_20 = Nothing ' -- 초기화 

Class C_TB_20
	' -- 테이블의 컬럼변수 
	Dim W1
	Dim W2
	Dim W3
	Dim W4
	Dim W5
	Dim W6
	
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.

	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
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
	End Function
	
	Function EOF()
		EOF = lgoRs1.EOF
	End Function
	
	Function MoveFist()
		lgoRs1.MoveFirst
		Call GetData()
	End Function
	
	Function MoveNext()
		lgoRs1.MoveNext
		Call GetData()
	End Function	
	
	Function GetData()
		If Not lgoRs1.EOF Then
			W1			= lgoRs1("W1")
			W2			= lgoRs1("W2")
			W3			= lgoRs1("W3")
			W4			= lgoRs1("W4")
			W5			= lgoRs1("W5")
			W6			= lgoRs1("W6")
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
				lgStrSQL = lgStrSQL & " FROM TB_20	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W3129MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W3129MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, sT1_400SUM, sT2_100Sum
    Dim chkW1_1,chkW1_2,chkW1_3, chkW1_4, chkW1_5, chkW1_6, chkW1_7,chkData,chkW1
    
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W3129MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W3129MA1"

	Set lgcTB_20 = New C_TB_20		' -- 해당서식 클래스 
	
	If Not lgcTB_20.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W3129MA1

	'==========================================
	' -- 제20호 감가상각비명세서 합계표 오류검증 
	iSeqNo = 1	: sHTFBody = ""
	 
	 chkW1_1 = "기말현재액"
	 chkW1_2 = "감가상각누계액"
	 chkW1_3 = "미상각잔액"
	 chkW1_4 = "상각범위액"
	 chkW1_5 = "회사손금계상액"
	 chkW1_6 = "상각부인액"
	 chkW1_7 = "시인부족액"

	
	Do Until lgcTB_20.EOF 
	
		sHTFBody = sHTFBody & "83"
		sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
		If Not ChkNotNull(lgcTB_20.W1, "자산구분코드") Then blnError = True	
		chkW1 = ""
		sTmp = lgcTB_20.W1
		Select Case sTmp
			Case "1"
				 chkW1 = chkW1_1         '메세지에서 구분코드로 넣지 않고 말로 넣기 위해 
			     chkW1_1 = ""
			Case "2"
				 chkW1 = chkW1_2
			     chkW1_2 = ""
			Case "3"
				 chkW1 = chkW1_3
			     chkW1_3 = ""
			Case "4"
				 chkW1 = chkW1_4
				 chkW1_4 = ""
			Case "5"
				 chkW1 = chkW1_5
				 chkW1_5 = ""
			Case "6"
				 chkW1 = chkW1_6
				 chkW1_6 = ""
			Case "7"
				 chkW1 = chkW1_7	
			     chkW1_7 = ""
		End Select
       
		
		sHTFBody = sHTFBody & UNIChar(lgcTB_20.W1, 1)
		
		
		
		'②합계액 : ③건출물 + ④기계장치 + ⑤기타자산 + ⑥무형고정자산 
		if  UNICDbl(lgcTB_20.W2, 0) <>  UNICDbl(lgcTB_20.W3,0) +  UNICDbl(lgcTB_20.W4, 0) +  UNICDbl(lgcTB_20.W5, 0) + UNICDbl(lgcTB_20.W6,  0) then
		   	Call SaveHTFError(lgsPGM_ID, lgcTB_20.W2, UNIGetMesg(TYPE_CHK_NOT_EQUAL,chkW1 & "의 합계액","(3)건축물 + (4)기계장치 + (5)기타자산 + (6)무형고정자산 "))
		   	 blnError = True	
		end if
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W2, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W3, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W4, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W5, 15, 0)
		sHTFBody = sHTFBody & UNINumeric(lgcTB_20.W6, 15, 0)
		
		sHTFBody = sHTFBody & UNIChar("", 18) & vbCrLf	' -- 공란 
	  
		lgcTB_20.MoveNext 
	Loop
	
	
	'자산구분코드가 레코드 별로 모두 들어와야 한다.
	'( 레코드가 7개가 생성 되어야 한다.)

	if chkW1_1 <> "" Or chkW1_2 <> "" or chkW1_3<>"" Or chkW1_4<>"" Or chkW1_5<>"" Or chkW1_6<>""  Or chkW1_7<>"" then
	    chkData =chkW1_1 & " " & chkW1_2  &  " " & chkW1_3 &  " " & chkW1_4 & " " & chkW1_5 &  " " & chkW1_6 &  " " & chkW1_7
	   	Call SaveHTFError(lgsPGM_ID,Trim(chkData), UNIGetMesg(TYPE_CHK_NULL, chkData))
	   	 blnError = True	
	end if  
	
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_20 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W3129MA1(pMode, pCode1, pCode2, pCode3)
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
	PrintLog "SubMakeSQLStatements_W3129MA1 : " & lgStrSQL
End Sub
%>
