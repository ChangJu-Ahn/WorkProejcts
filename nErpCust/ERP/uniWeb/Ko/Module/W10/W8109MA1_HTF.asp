<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        :  제13호 농어촌특별과세 
'*  3. Program ID           : W8109MA1
'*  4. Program Name         : W8109MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_13

Set lgcTB_13 = Nothing ' -- 초기화 

Class C_TB_13
	' -- 테이블의 컬럼변수 
	Dim WHERE_SQL		' -- 기본 검색조건(법인/사업연도/신고구분)외의 검색조건 
	Dim	CALLED_OUT		' -- 외부에서 부른 경우 
	
	Private lgoRs1		' -- 멀티로우 데이타는 지역변수로 선언한다.
	Private lgoRs2		' -- 멀티로우 데이타는 지역변수로 선언한다.

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
		Call MoveFirst(pType)
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
	
	Function MoveFirst(Byval pType)
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
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_13_A	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "H2"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_13_B	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W8109MA1
	Dim A106

End Class

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W8109MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, sMsg ,dblSum,dblW10,strType,strMsg
    Dim iSeqNo
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W8109MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W8109MA1"

	Set lgcTB_13 = New C_TB_13		' -- 해당서식 클래스 
	
	If Not lgcTB_13.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
	' -- 참조할 서식 선언 


	'==========================================
	' --  제13호 농어촌특별과세 오류검증 
	' -- 1. 매출및매입거래등 
	
	sHTFBody = sHTFBody & "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
    dblSum = 0
    
    
    lgcTB_13.Find 1, "W2_CD='101'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='102'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='103'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='104'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='111'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='112'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='113'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='000'"	' -- (7)비과세 
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='121'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='122'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='126'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='127'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='128'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='129'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='125'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='131'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='140'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='132'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='133'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='134'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='135'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='136'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='137'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='138'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='141'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='142'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W2_CD='139'"
    sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    lgcTB_13.Find 1, "W1_CD='10'"
    sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
    
    ' -- 파일생성 순서가 변경됨 
    ' ------------------------------------------------------------
    lgcTB_13.Find 1, "W2_CD='101'"
    
	Do Until lgcTB_13.EOF(1) 
		
		

		If lgcTB_13.GetData(1, "W1_CD")  ="07" then
		   dblSum = dblSum + UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
		Elseif  lgcTB_13.GetData(1, "W1_CD")  ="10"    then
			dblW10 = UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
		End if
		
		
       sTmp = lgcTB_13.GetData(1, "W2_CD")
		'항목 (7) + (121) + (122) + (123) + (124) + (125) + (131) + (132) + (133)+ (134) + (135) + (136) + (137) + (138) + (139)
		Select Case sTmp   '합을 구하기 위해서 
			Case "121" , "122" , "123" , "124" , "125" , "131" , "132" , "133", "134" , "135" , "136" , "137" , "138" , "139"
				  dblSum = dblSum + UNIcdbl(lgcTB_13.GetData(1, "W4") ,0)
	    End Select
		
		
		Select Case sTmp
			Case "104", "113", "125", "139"
				'sHTFBody = sHTFBody & UNIChar(lgcTB_13.GetData(1, "W2"), 40)
				
				If Not ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_감면세액") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
				
			Case "121"	,"122" , "131" , "132" , "134" , "135", "136", "138"
						Select Case sTmp
						       Case "121" 
									  strType = "A106_W23"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(23) 공제세액"
						       Case "122" 
									  strType = "A106_W03"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(03)공제세액"
						       Case "123" 
									  strType = "A106_W14"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(14)공제세액"
						       Case "131" 
									  strType = "A106_W31"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(31)공제세액"
						        Case "132" 
									  strType = "A106_W75"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(75)공제세액"
						       Case "133" 
									  strType = "A106_W78"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(78)공제세액"      
						        Case "134" 
									  strType = "A106_W35"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(35)공제세액"       
						                                   
						       Case "135" 
									  strType = "A106_W36"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(36)공제세액"    
						                                       
								 Case "136" 
									  strType = "A106_W77"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(77)공제세액" 
						            
						       Case "138" 
									  strType = "A106_W42"
						              strMsg = " 공제감면세액및추가납부세액합계표(갑)(A106)의 코드(42)공제세액"      
						End Select 
			
			      If  ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_감면세액") Then 
			      
			 
			          if unicdbl(lgcTB_13.GetData(1, "W4"),0) <> unicdbl(Getdata_TB_3_13ho(strType) ,0) then
			             Call SaveHTFError(lgsPGM_ID, lgcTB_13.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, lgcTB_13.GetData(1, "W2") & "_감면세액",strMsg))
			             blnError = True	
			          End if   
			      Else
			      
			          blnError = True	
			      End if    
			    	'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
			      
			Case Else
				If Not ChkNotNull(lgcTB_13.GetData(1, "W4"), lgcTB_13.GetData(1, "W2") & "_감면세액") Then blnError = True	
				'sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(1, "W4"), 15, 0)
		
		End Select

		Call lgcTB_13.MoveNext(1)	' -- 1번 레코드셋 
	Loop

    if UNICDbl(dblSum,0) <> Unicdbl(dblW10,0)  then
       '감면세액합계 = 항목 (7) + (121) + (122) + (123) + (124) + (125) + (131) + (132) + (133)+ (134) + (135) + (136) + (137) + (138) + (139)
        Call SaveHTFError(lgsPGM_ID, dblSum, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "감면세액합계","각 항목들의 합"))
        blnError = True	
    End if

	If Not lgcTB_13.EOF(2) Then
		
		If Not ChkNotNull(lgcTB_13.GetData(2, "W1"), "법인세과세표준") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, Replace("W2_VAL","%","")), "조세특례제한법 제72조 세율") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W3"), "산출세액") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W4"), "과세표준_금액") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W6"), "산출세액") Then blnError = True	
		If Not ChkNotNull(lgcTB_13.GetData(2, "W7"), "감면세액") Then blnError = True	
	
	End If

	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W1"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, Replace("W2_VAL","%","")), 5, 2)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W3"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W4"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W6"), 15, 0)
	sHTFBody = sHTFBody & UNINumeric(lgcTB_13.GetData(2, "W7"), 15, 0)
	
	sHTFBody = sHTFBody & UNIChar("", 34) & vbCrLf	' -- 공란 
	
	' ----------- 
	'Call SubCloseRs(oRs2)
	
	PrintLog "Write2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call Write2File(sHTFBody)
	End If
	

	Set lgcTB_13 = Nothing	' -- 메모리해제 
	
End Function


Function Getdata_TB_3_13ho(byval pType)
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- 제8호 갑 공제감면세액명세서 
        Set cDataExists = new TYPE_DATA_EXIST_W8109MA1
		Set cDataExists.A106  = new C_TB_8A	' -- W8107MA1_HTF.asp 에 정의됨 
		cDataExists.A106.CALLED_OUT	= True		' -- 외부에서 호출함을 알림 
		cDataExists.A106.WHERE_SQL = ""			' -- 클래스가 기본적으로 로드하는 조건외의 추가 조건을 지							
		' -- 추가 조회조건을 읽어온다.
		
		Call SubMakeSQLStatements_W8109MA1(pType,iKey1, iKey2, iKey3)   
	        cDataExists.A106.WHERE_SQL = lgStrSQL
		
		If Not cDataExists.A106.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제8호 (갑)공제감면세액서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
		
			 dblData = UNICDbl(cDataExists.A106.w4,0)
					
	       	
		End If	
						
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A106 = Nothing
		Set cDataExists = Nothing	' -- 메모리해제 
	
		Getdata_TB_3_13ho = unicdbl(dblData,0)


End Function


' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W8109MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A106_W23" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '23'" 	 & vbCrLf
	  Case "A106_W03" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '03'" 	 & vbCrLf		
	  Case "A106_W14" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '14'" 	 & vbCrLf		
	  Case "A106_W31" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '31'" 	 & vbCrLf		
	  Case "A106_W75" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '75'" 	 & vbCrLf		
	  Case "A106_W78" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '78'" 	 & vbCrLf	
			
	  Case "A106_W35" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '35'" 	 & vbCrLf		
     
     Case "A106_W36" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '36'" 	 & vbCrLf																						

	  Case "A106_W77" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '77'" 	 & vbCrLf			
			
	  Case "A106_W42" '-- 외부 참조 SQL
           
			lgStrSQL = ""
	
			lgStrSQL = lgStrSQL & "	AND A.w2_1	= '42'" 	 & vbCrLf																						
																					

	End Select
	PrintLog "SubMakeSQLStatements_W8109MA1 : " & lgStrSQL
End Sub
%>
