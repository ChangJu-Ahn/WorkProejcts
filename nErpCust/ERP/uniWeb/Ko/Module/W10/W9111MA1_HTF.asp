<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제54호 주식등 변동상황명세서(갑)
'*  3. Program ID           : W9111MA1
'*  4. Program Name         : W9111MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_54_BP

Set lgcTB_54_BP = Nothing ' -- 초기화 

Class C_TB_54_BP
	' -- 테이블의 컬럼변수 

	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs3		' -- 멀티로우 데이타는 지역변수로 선언한다.
	
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	Function LoadData()
		Dim iKey1, iKey2, iKey3, blnData1, blnData2
				 
		'On Error Resume Next                                                             '☜: Protect system from crashing
		Err.Clear     
		
		LoadData = False
		blnData1 = True : blnData2 = True
		
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
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs2,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData2 = False
		End If
		
				
		If blnData1 = False And blnData2 = False Then
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
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 2
				lgoRs2.Filter = pWhereSQL
		End Select
	End Function
	
	Function EOF(Byval pType)
		Select Case pType
			Case 1
				EOF = lgoRs1.EOF
			Case 2
				EOF = lgoRs2.EOF
		End Select
	End Function
	
	Function MoveFist(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveFirst
			Case 2
				lgoRs2.MoveFirst
		End Select
	End Function
	
	Function MoveNext(Byval pType)
		Select Case pType
			Case 1
				lgoRs1.MoveNext
			Case 2
				lgoRs2.MoveNext
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
		End Select
	End Function
	
	Function RecordCount(Byval pType)
		Select Case pType
			Case 2
				RecordCount = lgoRs2.RecordCount
		End Select
	End Function

	Function CloseRs()	' -- 외부에서 닫기 
		Call SubCloseRs(lgoRs1)
		Call SubCloseRs(lgoRs2)
	End Function
		
	'----------- 클래스 시작/종료 이벤트 -------------
	Private Sub Class_Initialize()
		CALLED_OUT = False
	End Sub

	Private Sub Class_Terminate()
		Call SubCloseRs(lgoRs1)		' -- 레코드셋이 지역(전역)이므로 클래스 파괴시에 해제한다.
		Call SubCloseRs(lgoRs2)	
	End Sub

	' ------------------ 조회 SQL 함수 --------------------------------
	Private Sub SubMakeSQLStatements(pMode, pCode1, pCode2, pCode3)
	    Select Case pMode 
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.* , " & vbCrLf
	            lgStrSQL = lgStrSQL & " B.TAX_OFFICE, B.FISC_START_DT, B.FISC_END_DT " & vbCrLf
	            lgStrSQL = lgStrSQL & " , B.OWN_RGST_NO, B.CO_NM, B.REPRE_NM " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54_BPH	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & "		INNER JOIN TB_COMPANY_HISTORY	B  WITH (NOLOCK) ON A.CO_CD=B.CO_CD AND A.FISC_YEAR=B.FISC_YEAR AND A.REP_TYPE=B.REP_TYPE " & vbCrLf	' 서식3호 

				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_54_BPD	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W9111MA1
	Dim A103

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W9111MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, arrTmp1(1), arrTmp2(1)
    Dim iSeqNo, iDx, iDtlCnt
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W9111MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W9111MA1"

	Set lgcTB_54_BP = New C_TB_54_BP		' -- 해당서식 클래스 
	
	If Not lgcTB_54_BP.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W9111MA1

	'==========================================
	' -- 제54호부표주식출자지분양도명세서 오류검증 
	' -- 1. 기본사항(83A132)
	sHTFBody = "B"
	sHTFBody = sHTFBody & UNIChar("41", 2)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "TAX_OFFICE"), "세무서코드") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "TAX_OFFICE"), 3)
				
	If  ChkNotNull(lgcTB_54_BP.GetData(1, "OWN_RGST_NO"), "신고법인_사업자등록번호") Then 
	    IF UNIRemoveDash(lgcTB_54_BP.GetData(1, "OWN_RGST_NO")) <>  UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO) Then
	       '신고법인의 사업자등록번호	- 법인세신고공통(A100)의 납세자ID와 일치 
	        Call SaveHTFError(lgsPGM_ID, lgcTB_54_BP.GetData(1, "OWN_RGST_NO"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "신고법인의 사업자등록번호", "법인세신고공통(A100)의 납세자ID"))
	        blnError = True	
	    End If 
	Else
	    blnError = True	
	End If    
	sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(1, "OWN_RGST_NO")), 10)
		
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "CO_NM"), "신고법인_법인명") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "CO_NM"), 40)
		
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "REPRE_NM"), "신고법인_대표자성명") Then blnError = True	
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "REPRE_NM"), 30)
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "FISC_START_DT"), "사업연도(시작일자)") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(1, "FISC_START_DT"))
	
	If Not ChkNotNull(lgcTB_54_BP.GetData(1, "FISC_END_DT"), "사업연도(종료일자)") Then blnError = True	
	sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(1, "FISC_END_DT"))
		
	If Not ChkBoundary("1,2,3,4,5,6,7", lgcTB_54_BP.GetData(1, "W4"), "주식구분") Then blnError = True
	sHTFBody = sHTFBody & UNIChar(lgcTB_54_BP.GetData(1, "W4"), 1)
	
	sHTFBody = sHTFBody & UNINumeric(UNICDBL(lgcTB_54_BP.RecordCount(2),0), 6, 0)	' 그리드 갯수 
		
	sHTFBody = sHTFBody & UNIChar("", 91)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	blnError = False : sHTFBody = ""
		
	' -- 주식수변동상황 
	iSeqNo = 1	
	
	Do Until lgcTB_54_BP.EOF(2) 

		sHTFBody = sHTFBody & "C"
		sHTFBody = sHTFBody & UNIChar("41", 2)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		sHTFBody = sHTFBody & UNIChar(lgcCompanyInfo.TAX_OFFICE, 3)	' -- 세무서코드 
		sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcCompanyInfo.OWN_RGST_NO), 10) ' -- 사업자번호 
	    
	    
	    If iSeqNo <> 1 Then
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W6"), "양도자_주민(사업자)등록번호") Then blnError = True	
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W5"), "양도자_성명(법인명)") Then blnError = True	
		   If Not ChkNotNull(lgcTB_54_BP.GetData(2, "W9"), "양도주식수(출자좌수)") Then blnError = True	
		 End If  
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(2, "W6")), 13)
		'양도자_주민(사업자)등록번호	- 양도자_주민(사업자)등록번호가 ‘’또는 ‘'AAAAAAAAAAAAA'’인 경우는 일련번호가 ‘000001’ 이 아니면 오류 
		' 주민번호는 ‘BBBBBBBBBBBBB’이 아니거나 ‘CCCCCCCCCC’,’ DDDDDDDDDD’,’’ EEEEEEEEEE’,’ FFFFFFFFFF’으로 시작하지 않고 또는 앞의 4자리가 0000이거나 4자리 이하이면 오류 
		
		sHTFBody = sHTFBody & UNIChar(UNIRemoveDash(lgcTB_54_BP.GetData(2, "W5")), 40)
		
		If lgcTB_54_BP.GetData(2, "W7") <> "" Then 
		   sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(2, "W7"))
		Else
		   sHTFBody = sHTFBody & UNIChar( "", 8) 
		End IF   
		
		
		If  lgcTB_54_BP.GetData(2, "W8") <> "" Then 
		   sHTFBody = sHTFBody & UNI8Date(lgcTB_54_BP.GetData(2, "W8"))
		Else
		   sHTFBody = sHTFBody & UNIChar("", 8) 
		End IF   
		   
		
		
		sHTFBody = sHTFBody & UNINumeric(lgcTB_54_BP.GetData(2, "W9"), 13, 0)	
		
		sHTFBody = sHTFBody & UNIChar("", 96) & vbCrLf

		iSeqNo = iSeqNo + 1
		
		Call lgcTB_54_BP.MoveNext(2)	' -- 1번 레코드셋 
	Loop

	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
		
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_54_BP = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W9111MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W9111MA1 : " & lgStrSQL
End Sub
%>
