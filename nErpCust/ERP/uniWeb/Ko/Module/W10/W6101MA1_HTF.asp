<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제48호 소득구분계산서 
'*  3. Program ID           : W6101MA1
'*  4. Program Name         : W6101MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_48

Set lgcTB_48 = Nothing ' -- 초기화 

Class C_TB_48
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.
	
	' ------------------ 클래스 데이타 로드 함수 --------------------------------
	
	
	
	Function Clone(Byref pRs)
		Set pRs = lgoRs2.clone
	End Function



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
		Call SubMakeSQLStatements("H1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		' -- 커서를 클라이언트로 변경 **주의 ../wcm/incServerADoDb.asp 에만 지원되는 기능 
		gCursorLocation = adUseClient 

		If   FncOpenRs("P",lgObjConn,lgoRs1,lgStrSQL, adOpenKeySet, adLockReadOnly) = False Then
			blnData1 = False
		End If
		
		' --서식을 읽어온다.
		Call SubMakeSQLStatements("H2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

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
			Case 1
				lgoRs1.Find pWhereSQL
			Case 2
				lgoRs2.Find pWhereSQL
		End Select
	End Function

	Function Filter(Byval pType, Byval pWhereSQL)
		Select Case pType
			Case 1
				lgoRs1.Filter = pWhereSQL
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
			Case "H1"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_48H1	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
			
			Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_48H2	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
			
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6101MA1
	Dim A103

End Class




' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6101MA1()
    Dim iKey1, iKey2, iKey3
    Dim arrHTFBody(1), blnError, oRs2, sTmp, cDataExists
    Dim iSeqNo, arrGamMyun(5), i, iPageNo
    Dim sAmt1,sAmt2, sAmt3
    Dim sRate1,sRate2, sRate3
    Dim dblAmt1
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6101MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6101MA1"

	Set lgcTB_48 = New C_TB_48		' -- 해당서식 클래스 
	
	If Not lgcTB_48.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 
	Set cDataExists = new TYPE_DATA_EXIST_W6101MA1

	'==========================================
	' -- 제48호 소득구분계산서 오류검증 
	' -- 1. 수입조정계산 
	iSeqNo = 0	
    	Call lgcTB_48.Clone(oRs2)	' 서식검증에 필요한 참조 레코드셋을 복제 
	' 감면사업등록 처리 부터 
	Do Until lgcTB_48.EOF(1) 
	
		If UNICDbl(lgcTB_48.GetData(1, "W_TYPE") , 0) < 4 Or (UNICDbl(lgcTB_48.GetData(1, "W_TYPE") , 0) >= 4 and Trim(lgcTB_48.GetData(1, "W_NM"))="") Then
			iPageNo = 0		' 페이지번호 
		Else
		    iPageNo = 1		' 페이지번호 
		End If
		arrGamMyun(iSeqNo) = "" & lgcTB_48.GetData(1, "W_NM")
		iSeqNo = iSeqNo + 1
		Call lgcTB_48.MoveNext(1)	' -- 2번 레코드셋 
	Loop

	Do Until lgcTB_48.EOF(2) 
		iSeqNo = 0
		
		For i = 0 To iPageNo		' 페이지 만큼 루프\
		   if i = 0 then
		      sAmt1		= "W4_1"
		      sRate1	= "W5_1"
		      sAmt2		= "W4_2"
		      sRate2	= "W5_2"
			  sAmt3		= "W4_3"
		      sRate3	= "W5_3"
		   Elseif i = 1 then
		      sAmt1		= "W4_4"
		      sRate1	= "W5_4"
		      sAmt2		= "W4_5"
		      sRate2	= "W5_5"
			  sAmt3		= "W4_6"
		      sRate3	= "W5_6"
		   
		   end if   
			arrHTFBody(i) = arrHTFBody(i) & "83"
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
		
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(i+1, 2, 0)		' 페이지번호 
		
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo), 50)	' 감면사업 
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo+1), 50)	' 감면사업 
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(arrGamMyun(iSeqNo+2), 50)	' 감면사업 
		
			If  ChkNotNull(lgcTB_48.GetData(2, "W1_CD"), "과목구분코드") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNIChar(UNICDbl(lgcTB_48.GetData(2, "W1_CD"), 0), 2)		' 디비엔 01, 홈텍스엔 1로 정의되어 UNICDbl 해 
		
			If  ChkNotNull(lgcTB_48.GetData(2, "W3"), "합계") Then
			  
			    if iPageNo = i Then   '합계 체크이기때문에 한번만 체크해주면 된다.
						Stmp = UNICDbl(lgcTB_48.GetData(2, "W4_1"),0) + UNICDbl(lgcTB_48.GetData(2, "W4_2"),0) +  UNICDbl(lgcTB_48.GetData(2, "W4_3"),0) _
						       +  UNICDbl(lgcTB_48.GetData(2, "W4_4"),0)+ UNICDbl(lgcTB_48.GetData(2, "W4_5"),0) + UNICDbl(lgcTB_48.GetData(2, "W4_6"),0)+ UNICDbl(lgcTB_48.GetData(2, "W6"),0)
						      
						If Unicdbl(lgcTB_48.GetData(2, "W3"),0) <> Stmp Then 
						
						   Call SaveHTFError(lgsPGM_ID,lgcTB_48.GetData(2, "W3"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드" & lgcTB_48.GetData(2, "W1_CD")  & "합계","각 감면분 금액의 합"))
						   blnError = True	
						End if
						
			    
			    
						oRs2.MoveFirst	
						if lgcTB_48.GetData(2, "W1_CD") = "03" Then '(3) 매출 총이익 : 매출액(1) - 매출원가(2)
						   oRs2.Find  "W1_CD = '01'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '02'"
						   dblAmt1 = dblAmt1 -  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "매출 총이익","매출액(1) - 매출원가(2)"))
							  blnError = True	
							End If	
						End if  
			    
						if lgcTB_48.GetData(2, "W1_CD") = "06" Then '(6) 판매비와 관리비 계 : 개별분(4) + 공통분(5)
						   
						   oRs2.Find  "W1_CD = '04'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '05'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "판매비와 관리비 계 "," 개별분(4) + 공통분(5)"))
							  blnError = True	
							End If	
						End if   
			    
						if lgcTB_48.GetData(2, "W1_CD") = "07" Then '(7) 영업이익 : 매출총이익(3) -판매비와 관리비 계(6)
						   oRs2.MoveFirst
						   oRs2.Find  "W1_CD = '03'"
						   dblAmt1 =   Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '06'"
						   dblAmt1 = dblAmt1 -   Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "영업이익 "," 매출총이익(3) -판매비와 관리비 계(6)"))
							  blnError = True	
							End If	
						End if   
			    
						if lgcTB_48.GetData(2, "W1_CD") = "10" Then '(10) 영업외수익 계 : 개별분(8) + 공통분(9)
						   oRs2.Find  "W1_CD = '08'"
						   dblAmt1 =   Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '09'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, " 영업외수익 계 ","개별분(8) + 공통분(9)"))
							  blnError = True	
							End If	
						End if   
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "13" Then '(13) 영업외비용 계 : 개별분(11) + 공통분(12)
						   oRs2.Find  "W1_CD = '11'"
						   dblAmt1 =  Unicdbl(oRs2("W3"),0)
						   oRs2.Find  "W1_CD = '12'"
						   dblAmt1 = dblAmt1 +  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  영업외비용 계 ","개별분(11) + 공통분(12)"))
							  blnError = True	
							End If	
						End if    
			    
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "14" Then '(14) 경상이익 : 영업이익(7) + 영업외수익 계(10) - 영업외비용 계(13)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '07'"
						    dblAmt1 =   Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '10'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '13'"
						    dblAmt1 = dblAmt1 -  Unicdbl(oRs2("W3"),0)
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  경상이익 ","영업이익(7) + 영업외수익 계(10) - 영업외비용 계(13)"))
							  blnError = True	
							End If	
						End if    
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "17" Then '(17) 특별이익 계 : 개별분(15) + 공통분(16)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '15'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '16'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  특별이익 계 ","개별분(15) + 공통분(16)"))
							  blnError = True	
							End If	
						End if  
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "20" Then '(20) 특별손실 계 : 개별분(18) + 공통분(19)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '18'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '19'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  특별손실 계 ","개별분(18) + 공통분(19)"))
							  blnError = True	
							End If	
						End if      
			    
						  If lgcTB_48.GetData(2, "W1_CD") = "21" Then '(21) 각 사업연도소득 또는 설정 전 소득 : 경상이익(14) + 특별이익 계(17) - 특별손실 계(20)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '14'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '17'"
						    dblAmt1 = dblAmt1 + Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '20'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  각 사업연도소득 또는 설정 전 소득 ","경상이익(14) + 특별이익 계(17) - 특별손실 계(20)"))
							  blnError = True	
							End If	
						End if      
			    
						 If lgcTB_48.GetData(2, "W1_CD") = "25" Then '(25) 과세표준 : 각사업년도소득(21) - 이월결손금(22) - 비과세소득(23) - 소득공제액(24)
						    oRs2.MoveFirst
						    oRs2.Find  "W1_CD = '21'"
						    dblAmt1 =  Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '22'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '23'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
						    oRs2.Find  "W1_CD = '24'"
						    dblAmt1 = dblAmt1 - Unicdbl(oRs2("W3"),0)
			 
						   If UNICDbl(lgcTB_48.GetData(2, "W3"), 0) <> dblAmt1 Then
							  Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_48.GetData(2, "W3"), 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "  (25) 과세표준","각사업년도소득(21) - 이월결손금(22) - 비과세소득(23) - 소득공제액(24)"))
							  blnError = True	
							End If	
						End if      
			    End if
			Else
			    blnError = True	
			End If    
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W3"), 15, 0)
		
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt1), "감면분 등 금액_1") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt1), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate1), "감면분 등 비율_1") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate1), 5, 2)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt2), "감면분 등 금액_2") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt2), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate2), "감면분 등 비율_2") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate2), 5, 2)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sAmt3), "감면분 등 금액_3") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sAmt3), 15, 0)
			
			If Not ChkNotNull(lgcTB_48.GetData(2, sRate3), "감면분 등 비율_3") Then blnError = True	
			arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, sRate3), 5, 2)
		    
		    If iPageNo = 0  then
		        If Not ChkNotNull(lgcTB_48.GetData(2, "W6"), "기타분 금액") Then blnError = True	
			     arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W6"), 15, 0)
			    If Not ChkNotNull(lgcTB_48.GetData(2, "W7"), "기타분 비율") Then blnError = True	
			    arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W7"), 5, 2)
			    
		    Elseif iPageNo = 1 and i = iPageNo then
		        If Not ChkNotNull(lgcTB_48.GetData(2, "W6"), "기타분 금액") Then blnError = True	
			     arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W6"), 15, 0)
			    If Not ChkNotNull(lgcTB_48.GetData(2, "W7"), "기타분 비율") Then blnError = True	
			    arrHTFBody(i) = arrHTFBody(i) & UNINumeric(lgcTB_48.GetData(2, "W7"), 5, 2)
			    
			Else
			    arrHTFBody(i) = arrHTFBody(i) &  UNINumeric("0", 15, 0)
			    arrHTFBody(i) = arrHTFBody(i) &  UNINumeric("0", 5, 2)  
		    End if
		    
		    
			If i = 0 Then
				arrHTFBody(i) = arrHTFBody(i) & UNIChar(lgcTB_48.GetData(2, "DESC1"), 30)
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 15) & vbCrLf	' -- 공란 
			Else
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 30)
				arrHTFBody(i) = arrHTFBody(i) & UNIChar("", 15) & vbCrLf	' -- 공란 
			End If
			iSeqNo = 3
		
			
		Next
		Call lgcTB_48.MoveNext(2)	' -- 1번 레코드셋 
	Loop

	PrintLog "Write2File : " & arrHTFBody(0) 
	PrintLog "Write2File : " & arrHTFBody(1)
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(arrHTFBody(0))
		Call Write2File(arrHTFBody(1))
	End If
			
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_48 = Nothing	' -- 메모리해제 
	
End Function

' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6101MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W6101MA1 : " & lgStrSQL
End Sub
%>
