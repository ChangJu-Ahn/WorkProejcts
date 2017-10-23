<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제4호 최저한세조정계산서 
'*  3. Program ID           : W6127MA1
'*  4. Program Name         : W6127MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_4

Set lgcTB_4 = Nothing ' -- 초기화 

Class C_TB_4
	' -- 테이블의 컬럼변수 
	
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
	
	End Function	
	
	Function GetData(Byval pFieldNm)
		If Not lgoRs1.EOF Then
			GetData		= lgoRs1(pFieldNm)
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
				lgStrSQL = lgStrSQL & " FROM TB_4	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6127MA1
	Dim A103
	Dim A101	' -- 법인세과세표준세액조정계산서(A101)
	Dim A108	' -- 특별비용조정명세서(A108) - W4105MA1
End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6127MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrVal(5, 25)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6127MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6127MA1"

	Set lgcTB_4 = New C_TB_4		' -- 해당서식 클래스 
	
	If Not lgcTB_4.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W6127MA1

	' -- 쿼리변수 
	iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
	iKey2 = FilterVar(lgsFISC_YEAR,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgsREP_TYPE,"''", "S")		' 신고구분 

	'==========================================
	' -- 제4호 최저한세조정계산서 오류검증 
	iSeqNo = 1	: sHTFBody = ""

	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

	Do Until lgcTB_4.EOF 
		' -- 2006-01-05 : 200603 개정판 
		arrVal(2, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W2"), 0)
		arrVal(3, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W3"), 0)
		arrVal(4, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W4"), 0)
		arrVal(5, CInt(lgcTB_4.GetData("W1"))) = UNICDbl(lgcTB_4.GetData("W5"), 0)
	
'1		자료구분 
'2		서식코드 




		Select Case lgcTB_4.GetData("W1")
			Case "01"
			'3		결산서상당기순이익_감면후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "결산서상당기순이익_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "02"
			'4		익금산입_감면후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "익금산입_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "03"
			'5		손금산입_감면후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "손금산입_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
			Case "04"
			'6		조정후소득금액_감면후세액 
			'7		조정후소득금액_최저한세 
			'8		조정후소득금액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "조정후소득금액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "조정후소득금액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "조정후소득금액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "05"
			'9		준비금_최저한세 
			'10		준비금_조정감 
			'11		준비금_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "준비금_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "준비금_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "준비금_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "06"
			'12		특별상각_최저한세 
			'13		특별상각_조정감 
			'14		특별상각_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "특별상각_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "특별상각_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "특별상각_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "07"
			'15		특별비용손금산입전 소득금액_감면후세액 
			'16		특별비용손금산입전 소득금액_최저한세 
			'17		특별비용손금산입전 소득금액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "특별비용손금산입전소득금액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "특별비용손금산입전소득금액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "특별비용손금산입전소득금액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "08"
			'18		기부금한도초과액_감면후세액 
			'19		기부금한도초과액_최저한세 
			'20		기부금한도초과액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "기부금한도초과액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "기부금한도초과액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "기부금한도초과액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "09"
			
			'21		기부금한도초과이월액 손금산입_감면후세액 
			'22		기부금한도초과이월액 손금산입_최저한세 
			'23		기부금한도초과이월액 손금산입_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "기부금한도초과이월액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "기부금한도초과이월액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "기부금한도초과이월액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "10"
			'24		각사업년도소득금액_감면후세액 
			'25		각사업년도소득금액_최저한세 
			'26		각사업년도소득금액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "각사업연도소득금액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "각사업연도소득금액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "각사업연도소득금액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "11"
			'27		이월결손금_감면후세액 
			'28		이월결손금_최저한세 
			'29		이월결손금_조정후세액 


				If Not ChkNotNull(lgcTB_4.GetData("W2"), "이월결손금_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "이월결손금_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "이월결손금_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "12"
			'30		비과세소득_감면후세액 
			'31		비과세소득_최저한세 
			'32		비과세소득_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "비과세소득_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "비과세소득_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "비과세소득_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			
			Case "13"
			'33		최저한세적용대상비과세소득_최저한세  
			'34		최저한세적용대상비과세소득_조정감    
			'35		최저한세적용대상비과세소득_조정후세액 

				If Not ChkNotNull(lgcTB_4.GetData("W3"), "최저한세적용대상비과세소득_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "최저한세적용대상비과세소득_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "최저한세적용대상비과세소득_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "14"
			'36		최저한세적용대상익금불산입_최저한세  
			'37		최저한세적용대상익금불산입_조정감    
			'38		최저한세적용대상익금불산입_조정후세액			

				If Not ChkNotNull(lgcTB_4.GetData("W2"), "최저한세적용대상익금불산입_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "최저한세적용대상익금불산입_조정감 ") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "최저한세적용대상익금불산입_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "15"	
			'39		차가감소득금액_감면후세액            
			'40		차가감소득금액_최저한세              
			'41		차가감소득금액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "차가감소득금액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "차가감소득금액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "차가감소득금액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "16"

			'42		소득공제_감면후세액                  
			'43		소득공제_최저한세                    
			'44		소득공제_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "소득공제_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "소득공제_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "소득공제_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
				
			Case "17"
			'45		최저한세적용대상 소득공제_최저한세   
			'46		최저한세적용대상 소득공제_조정감     
			'47		최저한세적용대상 소득공제_조정후세액 		
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "최저한세적용대상소득공제_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "최저한세적용대상소득공제_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "최저한세적용대상소득공제_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "18"
			'48		과세표준금액_감면후세액              
			'49		과세표준금액_최저한세                
			'50		과세표준금액_조정후세액      
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "과세표준금액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "과세표준금액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "과세표준금액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "19"
			'51		세율_감면후세액                      
			'52		세율_최저한세                        
			'53		세율_조정후세액         
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "세율_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 5, 2)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "세율_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 5, 2)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "세율_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 5, 2)
				
			Case "20"
			'54		산출세액_감면후세액                  
			'55		산출세액_최저한세                    
			'56		산출세액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "산출세액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "산출세액_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"),  15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "산출세액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
				
			Case "21"
			'57		감면세액_감면후세액                  
			'58		감면세액_조정감                      
			'59		감면세액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "감면세액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "감면세액_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "감면세액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
		
			Case "22"
			'60		세액공제_감면후세액                  
			'61		세액공제_조정감                      
			'62		세액공제_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "세액공제_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W4"), "세액공제_조정감") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W4"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "세액공제_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)	
			Case "23"
			'63		차감세액_감면후세액                  
			'64		차감세액_조정후세액 
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "차감세액_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "차감세액_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "24"
			'65		선박표준이익_감면후세액              
			'66		선박표준이익_최저한세                
			'67		선박표준이익_조정후세액  
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "선박표준이익_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "선박표준이익_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "선박표준이익_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
			Case "25"
			'68		과세표준금액2_감면후세액             
			'69		과세표준금액2_최저한세               
			'70		과세표준금액2_조정후세액   
				If Not ChkNotNull(lgcTB_4.GetData("W2"), "과세표준금액2_감면후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W2"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W3"), "과세표준금액2_최저한세") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W3"), 15, 0)
				If Not ChkNotNull(lgcTB_4.GetData("W5"), "과세표준금액2_조정후세액") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_4.GetData("W5"), 15, 0)
		End Select
	
		lgcTB_4.MoveNext 
	Loop

	sHTFBody = sHTFBody & UNIChar("", 54) ' 200703 개정판 

	' -- 2006-01-05 : 200603 개정판 		
	' -- 제3호 법인세과세표준및세액조정계산서(A101)서식 
	Set cDataExists.A101 = new C_TB_3	' -- W8101MA1_HTF.asp 에 정의됨 
					
	' -- 추가 조회조건을 읽어온다.
	Call SubMakeSQLStatements_W6127MA1("A101",iKey1, iKey2, iKey3)   
					

	if 	arrVal(2, 23) = arrVal(3, 20)  Then	'같은경우 오류검증 안함!~200707	
	else
	'==============================================
	'오류검증 
	'==============================================
	
			If Not cDataExists.A101.LoadData() Then
				blnError = True
				Call SaveHTFError(lgsPGM_ID, "제3호 법인세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
			Else
	
				' -- 코드(01)결산서상당기순이익_감면후세액(W2)
				' 제3호 법인세과세표준및세액조정계산서(A101)의 코드(01) 결산서상당기순이익 과 일치 
				
				'call svrmsgbox (UNICDbl(cDataExists.A101.W01, 0) ,0,1)
				'call svrmsgbox (arrVal(2, 1) ,0,1)
				
				If UNICDbl(cDataExists.A101.W01, 0) <> arrVal(2, 1) Then	' -- W1: 01, W2의 값 
					blnError = True
					Call SaveHTFError(lgsPGM_ID, arrVal(2, 1), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "결산서상당기순이익_감면후세액","제3호 법인세과세표준및세액조정계산서(A101)의 코드(01) 결산서상당기순이익"))
				End If

				' -- 코드(02)익금산입_감면후세액 
				If arrVal(3, 5) > 0 Or arrVal(3, 6) > 0 Then
					' -- (05)준비금 또는 (06)특별상각의 (3)최저한세가 0 보다 크면 검증제외 
				Else
					' 제3호 법인세과세표준및세액조정계산서(A101)의 코드(02) 소득조정금액_익금산입 과 일치 
					If UNICDbl(cDataExists.A101.W02, 0) <> arrVal(2, 2) Then
						blnError = True
						Call SaveHTFError(lgsPGM_ID, arrVal(2, 2), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "익금산입_감면후세액","제3호 법인세과세표준및세액조정계산서(A101)의 코드(02) 소득조정금액_익금산입"))
					End If

					' -- 코드(03)손금산입_감면후세액(W2)
					' 제3호 법인세과세표준및세액조정계산서(A101)의 코드(03) 소득조정금액_손금산입 과 일치 
					If UNICDbl(cDataExists.A101.W03, 0) <> arrVal(2, 3) Then	
						blnError = True
						Call SaveHTFError(lgsPGM_ID, arrVal(2, 3), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "손금산입_감면후세액","제3호 법인세과세표준및세액조정계산서(A101)의 코드(03) 소득조정금액_손금산입"))
					End If
					
				End If

				
			End If

			' -- 사용한 클래스 메모리 해제	: 뒤에 또 사용됨 
			Set cDataExists.A101 = Nothing
			' -- 코드(05)준비금-최저한세(3) , 코드(06)특별상각_최저한세(3)  > 0 면, 특별비용조정명세서(A108) 필수 입력 

			If arrVal(3, 5) > 0 Or arrVal(3, 6) > 0 Then

				Set cDataExists.A108 = new C_TB_5	' -- W4105MA1_HTF.asp 에 정의됨 
								
				' -- 추가 조회조건을 읽어온다.
				Call SubMakeSQLStatements_W6127MA1("A108",iKey1, iKey2, iKey3)   
								
				cDataExists.A108.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
				cDataExists.A108.WHERE_SQL = lgStrSQL	' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지시함 
				'zzzzz 값 비교 안함?
				If Not cDataExists.A108.LoadData() Then
					blnError = True
					Call SaveHTFError(lgsPGM_ID, "제5호 특별비용조정명세서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
				End If
				
				Set cDataExists.A108 = Nothing
				
			End If

			' -- 200603 개정 : 선박표준이익이 0보다 클 경우 선박표준이익 산출명세서(A224) 항목(7) 선박표준이익과 일치 (미개발)
			If arrVal(5, 24) > 0 Then
	
			End If
	
						
			' 제3호 법인세과세표준및세액조정계산서(A101)의 코드(12) 산출세액 과 일치 
			'zzzz 확인묘망 
			'If UNICDbl(cDataExists.A101.W12, 0) <> arrVal(5, 20) Then	
			'	blnError = True
			'	Call SaveHTFError(lgsPGM_ID, arrVal(5, 20) & " <> " & cDataExists.A101.W12, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(20) 산출세액","제3호 법인세과세표준및세액조정계산서(A101)의 코드(12) 산출세액"))
			'End If


			' 코드(23) 항목(5) >= 코드(20) 항목(3)
			If arrVal(5, 23) >= arrVal(3, 20) Then	
			Else
				blnError = True
				Call SaveHTFError(lgsPGM_ID, arrVal(5, 20), "코드(23) 차감세액_조정후세액 은(는) 코드(20) 산출세액_최저한세 보다 같거나 커야 합니다.")
			End If

	end if				

	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_4 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6127MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A101" '-- 외부 참조 SQL

			lgStrSQL = ""

	End Select
	PrintLog "SubMakeSQLStatements_W6127MA1 : " & lgStrSQL
End Sub
%>
