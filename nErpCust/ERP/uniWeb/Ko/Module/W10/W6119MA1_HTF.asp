<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제8호 공제감몀세액(4)
'*  3. Program ID           : W6119MA1
'*  4. Program Name         : W6119MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
Dim lgcTB_8_4

Set lgcTB_8_4 = Nothing ' -- 초기화 

Class C_TB_8_4
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
	      Case "H"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_8_4H	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 

	      Case "D"
				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " SELECT  "
	            lgStrSQL = lgStrSQL & " A.*  " & vbCrLf
				lgStrSQL = lgStrSQL & " FROM TB_8_4D	A  WITH (NOLOCK) " & vbCrLf	' 서식3호 
				lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
				lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
				
				If WHERE_SQL <> "" Then lgStrSQL = lgStrSQL & WHERE_SQL		' -- 조회조건 추가 
		End Select
		PrintLog "SubMakeSQLStatements : " & lgStrSQL
	End Sub
	
End Class

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W6119MA1
	Dim A100
	Dim A101

End Class



  Function RtnQueryVal(strField,strFrom,strWhere)
        Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    RtnQueryVal = ""
	    Call CommonQueryRs(strField,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    RtnQueryVal = Replace(lgF0,Chr(11),"")
	    If RtnQueryVal = "X" Or trim(RtnQueryVal) = "" Or IsNull(RtnQueryVal) Then
                     
            ObjectContext.SetAbort
            Call SetErrorStatus
		End If
    End Function
    

' ------------------ 메인 함수 --------------------------------
Function MakeHTF_W6119MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists, dblW2W4
    Dim iSeqNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W6119MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W6119MA1"
    if Getdata_TB_1("A100")= "3" then
				Set lgcTB_8_4 = New C_TB_8_4		' -- 해당서식 클래스 
	
				If Not lgcTB_8_4.LoadData Then Exit Function			' -- 제3호2(1)(2)표준손익계산서 서식 로드 
	
				' -- 참조할 서식 선언 


				'==========================================
				' -- 제8호 공제감몀세액(4) 오류검증 
				' -- 1. 감면세액계산 
				sHTFBody = sHTFBody & "83"
				sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
					
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W1"), "구분") Then blnError = True	
				sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(1, "W1"), 20)
					
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W2"), "과세표준금액_감면대상사업") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W2"), 15, 0)
								
				If Not ChkNotNull(lgcTB_8_4.GetData(1, "W3"), "과세표준금액_비감면대상사업") Then blnError = True	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W3"), 15, 0)
					
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W4"), "과세표준금액_계") Then
					'과세표준금액_계= (2)감면대상사업금액 + (3)비감면사업금액 
					'- 법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준과 일치 
					If   unicdbl(lgcTB_8_4.GetData(1, "W4"),0) <> unicdbl(lgcTB_8_4.GetData(1, "W2"),0) + unicdbl(lgcTB_8_4.GetData(1, "W3"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준금액_계","(2)감면대상사업금액+ (3)비감면사업금액"))
					     blnError = True	
					End if
					
					
					If   unicdbl(lgcTB_8_4.GetData(1, "W4"),0) <> unicdbl(Getdata_TB_3("A101_10"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W4"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "과세표준금액_계","법인세과세표준및세액조정계산서(A101)의 코드(10)과세표준과 일치"))
					     blnError = True	
					End if
					
				Else
				     blnError = True	
				End if     	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W4"), 15, 0)
			
				
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W5"), "산출세액_계") Then
					'법인세과세표준및세액조정계산서(A101)의 코드(12)산출세액_계와 일치(A153의 감면대상금액이 “0”보다 큰 경우 반드시 입력 
					
					
					If   unicdbl(lgcTB_8_4.GetData(1, "W5"),0) <> unicdbl(Getdata_TB_3("A101_12"),0) then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W5"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "산출세액_계","법인세과세표준및세액조정계산서(A101)의 코드(12)산출세액_계"))
					     blnError = True	
					End if
					
				Else
				     blnError = True	
				End if     	
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W5"), 15, 0)
					
				If  ChkNotNull(lgcTB_8_4.GetData(1, "W6"), "감면사업대상소득비율") Then 
				    '항목 (2)감면대상사업금액 / (4)과세표준금액_계(단,100% 를 초과하는 경우, 100%로 기재)
				  
				    if unicdbl(lgcTB_8_4.GetData(1, "W4"),0)  <> 0 then
				        dblW2W4 =  (unicdbl(lgcTB_8_4.GetData(1, "W2"),0) / unicdbl(lgcTB_8_4.GetData(1, "W4"),0) ) 
				          if unicdbl(dblW2W4,0) * 100  > 100 then
					         dblW2W4 = unicdbl(dblW2W4,0) * 100
						  Else
					         dblW2W4 = unicdbl(dblW2W4,0)
					      end if
				      
				    Else
				        dblW2W4 = 0
				    End if    
				    
				    sTemp = UNICDbl(lgcTB_8_4.GetData(1, "W6"),0)/100   '%값으로 들어가있음 
				
					if unicdbl(sTemp,0) <> unicdbl(round(dblW2W4,2),0)  Then
					    Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W6"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "감면사업대상소득비율_계"," (2)감면대상사업금액 / (4)과세표준금액_계"))
					     blnError = True	
					   
					End if   
					
				    
				
				Else
				    blnError = True	
				End if
				
				sHTFBody = sHTFBody & UNINumeric(sTemp , 6, 3)
					
			If Not ChkNotNull(lgcTB_8_4.GetData(1, "W7"), "감면비율") Then blnError = True	
			   sTemp = UNICDbl(lgcTB_8_4.GetData(1, "W7"),0)/100   '%값으로 들어가있음 
			   sHTFBody = sHTFBody & UNINumeric(sTemp, 6, 3)
					
			   If  ChkNotNull(lgcTB_8_4.GetData(1, "W8"), "감면세액") Then
					    '항목 (5)산출세액 x (6)감면대상사업소득비율 x (7)감면비율(+100만 ~ -100만원 범위허용???)
					    dblW5W6W7 = unicdbl(lgcTB_8_4.GetData(1, "W5"),0) *  (unicdbl(lgcTB_8_4.GetData(1, "W6"),0)*0.01)  *  (unicdbl(lgcTB_8_4.GetData(1, "W7"),0) *0.01)
					  if   unicdbl(lgcTB_8_4.GetData(1, "W8"),0) <= unicdbl(dblW5W6W7,0) + 1000000 and unicdbl(lgcTB_8_4.GetData(1, "W8"),0) >= unicdbl(dblW5W6W7) - 1000000  then
					  Else
					        Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W8"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "감면세액"," (5)산출세액 x (6)감면대상사업소득비율 x (7)감면비율"))
					       blnError = True	
					  End if
				
				Else
	
				     blnError = True	
				End If     
				sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W8"), 15, 0)
	
				sHTFBody = sHTFBody & UNIChar("", 37) & vbCrLf	' -- 공란 

					
				Call lgcTB_8_4.MoveNext(1)	' -- 1번 레코드셋 


				PrintLog "WriteLine2File : " & sHTFBody
				' -- 파일에 기록한다.
				If Not blnError Then
					Call Write2File(sHTFBody)
				End If
	
				blnError = False : sHTFBody = ""
				' -- 2. 일반감면비율 
				iSeqNo = 1	

				' -- 데이타가 둘중에 한곳에만 들어감.
				Do Until lgcTB_8_4.EOF(2) 

					If lgcTB_8_4.GetData(2, "W_TYPE") = "1" Then	' 일반 
						sHTFBody = sHTFBody & "84"
					ElseIf lgcTB_8_4.GetData(2, "W_TYPE") = "2" Then ' 외국인 
						sHTFBody = sHTFBody & "85"
					End If
					sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 

					If UNICDbl(lgcTB_8_4.GetData(2, "SEQ_NO"), 0) <> 999999 Then
						sHTFBody = sHTFBody & UNISeqNo6(iSeqNo)
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W9"), "증자횟수") Then
						   If Not ChkBoundary("0,1,2,3,4" ,"증자횟수: " & lgcTB_8_4.GetData(2, "W9")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W10"), "구분") Then
						   If Not ChkBoundary("1,2,3,4,5,6" ,"구분: " & lgcTB_8_4.GetData(2, "W10")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
						
					
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W11"), "증자 등기 일자") Then blnError = True	
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W12"), "등록 일자") Then blnError = True	
					
					Else
						sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "SEQ_NO"), 6)
					End If
							
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W9"), 1)
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W10"), 1)
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W11"))
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W12"))
						
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W13"), "증자자본금_총액") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W13"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W14"), "증자자본금_외국투자가자본금") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W14"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W15"), "증자자본금_감면배제자본금") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W15"), 15, 0)
					
					If Not ChkDate(lgcTB_8_4.GetData(2, "W16_1"), "100% 감면기간 From") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W16_1"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W16_2"), "100% 감면기간 To") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W16_2"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W17_1"), "50% 감면기간 From") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W17_1"))
						
					If Not ChkDate(lgcTB_8_4.GetData(2, "W17_2"), "50% 감면기간 To") Then blnError = True	
					sHTFBody = sHTFBody & UNI8Date(lgcTB_8_4.GetData(2, "W17_2"))
						
				
					If UNICDbl(lgcTB_8_4.GetData(2, "SEQ_NO"), 0) <> 999999 Then
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W18"), "당해사업연도감면율") Then
						  If Not ChkBoundary("1,2,3" ,"당해사업연도감면율: " & lgcTB_8_4.GetData(2, "W18")  & " ") Then   blnError = True	
						Else
						    blnError = True	
						End if
				    End if		
						
							
					sHTFBody = sHTFBody & UNIChar(lgcTB_8_4.GetData(2, "W18"), 1)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W19"), "감면대상외국투자가자본금") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W19"), 15, 0)
					
					If Not ChkNotNull(lgcTB_8_4.GetData(2, "W20"), "당해사업연도총자본금") Then blnError = True	
					sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W20"), 15, 0)

					If lgcTB_8_4.GetData(2, "W_TYPE") = "1" Then	' 일반 
					
						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W21"), "감면비율") Then blnError = True	
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W21"), 6, 3)
						
						sHTFBody = sHTFBody & UNIChar("", 56) & vbCrLf	' -- 공란 
					ElseIf lgcTB_8_4.GetData(2, "W_TYPE") = "2" Then ' 외국인 

						If Not ChkNotNull(lgcTB_8_4.GetData(2, "W21"), "외국인투자비율") Then blnError = True	
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W21"), 6, 3)
						
						If  ChkNotNull(lgcTB_8_4.GetData(2, "W35"), "감면비율") Then 
						   '(항목(32)감면대상외국투자가자본금 / 항목(33)총감면대상외국투자가자본금 ) x 항목(34)외국인투자비율 
						    if unicdbl(lgcTB_8_4.GetData(1, "W20"),0) <> 0 then
						       if unicdbl(lgcTB_8_4.GetData(2, "W35"),0) <> (unicdbl(lgcTB_8_4.GetData(2, "W19"),0) /unicdbl(lgcTB_8_4.GetData(2, "W20"),0)) *unicdbl(lgcTB_8_4.GetData(1, "W21"),0) then
						           Call SaveHTFError(lgsPGM_ID, lgcTB_8_4.GetData(1, "W35"), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "감면대상외국투자가자본금_계","총감면대상외국투자가자본금 ) x 외국인투자비율"))
						           blnError = True	
						       End if 
						    End if   
						Else
						    blnError = True	
						End if    
						sHTFBody = sHTFBody & UNINumeric(lgcTB_8_4.GetData(1, "W35"), 6, 3)
						
						sHTFBody = sHTFBody & UNIChar("", 50) & vbCrLf	' -- 공란 
					End If

					iSeqNo = iSeqNo + 1
					
					Call lgcTB_8_4.MoveNext(2)	' -- 2번 레코드셋 
				Loop

				' ----------- 
				'Call SubCloseRs(oRs2)
	
				PrintLog "Write2File : " & sHTFBody
				' -- 파일에 기록한다.
				If Not blnError Then
					Call Write2File(sHTFBody)
				End If

				
				Set lgcTB_8_4 = Nothing	' -- 메모리해제 
		End if
End Function



Function Getdata_TB_1(byval strType)
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- 제8호 갑 공제감면세액명세서 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A100  = new C_TB_1	' -- W8107MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W6119MA1(strType,iKey1, iKey2, iKey3)   
				
		
				
		If Not cDataExists.A100.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제1호 법인세과세표준및세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else

			   dblData = UNICDbl(cDataExists.A100.W1, 0)

		End If	
						
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A100 = Nothing
		Set cDataExists = Nothing	' -- 메모리해제 
	
		Getdata_TB_1 = unicdbl(dblData,0)


End Function


Function Getdata_TB_3(byval strType )
 Dim dblData,iKey1, iKey2, iKey3,cDataExists

       ' -- 제8호 갑 공제감면세액명세서 
        Set cDataExists = new TYPE_DATA_EXIST_W6119MA1
		Set cDataExists.A101  = new C_TB_3	' -- W8101MA1_HTF.asp 에 정의됨 
							
		' -- 추가 조회조건을 읽어온다.
		Call SubMakeSQLStatements_W6119MA1(strType,iKey1, iKey2, iKey3)   
						

						
		If Not cDataExists.A101.LoadData() Then
			blnError = True
			Call SaveHTFError(lgsPGM_ID, " 제3호 법인세 과세표준 및 세액조정계산서", TYPE_DATA_NOT_FOUND)		' -- 외부에서 호출햇을땐 데이타없음을 저장해줘야 한다.
		Else
             Select Case  strType
               Case "A101_10"
			        'dblData = UNICDbl(cDataExists.A101.W10, 0)
			        dblData = UNICDbl(cDataExists.A101.W56, 0)	' 2006-01-05 (200603 개정판)
			   Case "A101_12"
			        dblData = UNICDbl(cDataExists.A101.W12, 0)
			 End Select  
		
		End If	
						
		' -- 사용한 클래스 메모리 해제 
		Set cDataExists.A101 = Nothing
		Set cDataExists = Nothing	' -- 메모리해제 
	
		Getdata_TB_3 = unicdbl(dblData,0)


End Function




' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W6119MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A100" '-- 외부 참조 SQL
	       lgStrSQL =""
      Case "A101" '-- 외부 참조 SQL
           lgStrSQL=""
	End Select
	PrintLog "SubMakeSQLStatements_W6119MA1 : " & lgStrSQL
End Sub
%>
