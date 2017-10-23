<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Function Name        : 제3호의3(3) 부속명세 분양원가 
'*  3. Program ID           : W1113MA1
'*  4. Program Name         : W1113MA1_HTF.asp
'*  5. Program Desc         : 전자신고 Conversion 프로그램 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

' ------------------ 전역 변수 --------------------------------
' -- 본 서식은 클래스를 W1107MA1_HTF 에 선언된걸 사용한다.

' -- 데이타 존재 체크 
Class TYPE_DATA_EXIST_W1113MA1
	Dim A115

End Class




' ------------------ 메인 함수 --------------------------------

Function Clone(Byref pRs)
	Set pRs = lgoRs1.clone   '복제 
End Function



Function MakeHTF_W1113MA1()
    Dim iKey1, iKey2, iKey3
    Dim sHTFBody, blnError, oRs2, sTmp, cDataExists
    Dim dblAmt1, dblAmt2, dblAmt3, arrNew(50)
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    blnError = False
    
    PrintLog "MakeHTF_W1113MA1 IS RUNNING: "
    
	lgsPGM_ID	= "W1113MA1"

	Set lgcTB_3_3_3 = New C_TB_3_3_3		' -- 해당서식 클래스 
	
	lgcTB_3_3_3.WHERE_SQL = "		AND A.W1 = '6' "		' 

	If Not lgcTB_3_3_3.LoadData Then Exit Function			
	
	
	'==========================================
	' -- 제3호의3(3) 부속명세 분양원가 전자신고 및 오류검증 
	sHTFBody = "83"
	sHTFBody = sHTFBody & UNIChar(lgsTAX_DOC_CD, 4)		' 특별한 변화가 없다면 호출프로그램에서 지정한 서식코드를 사용 
	
	
	Call lgcTB_3_3_3.Clone(oRs2)	' 서식검증에 필요한 참조 레코드셋을 복제 

	Do Until lgcTB_3_3_3.EOF 
	
		If  ChkNotNull(lgcTB_3_3_3.W5, lgcTB_3_3_3.W3) Then 
	    

					If lgcTB_3_3_3.W4 = "01" Then   '재료비 =코드 02 + 03
						oRs2.Find "W4 = '02'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						oRs2.Find "W4 = '03'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
					

						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2  Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "재료비","코드 02 + 03"))
						   blnError = True	
						End If

					End If
					
					
					If lgcTB_3_3_3.W4 = "03" Then   '(2)재료비 =코드 04 + 05 - 06
						oRs2.Find "W4 = '04'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						oRs2.Find "W4 = '05'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						oRs2.Find "W4 = '06'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
					
					
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 - dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "(2)재료비","코드 04 + 05 - 06"))
						   blnError = True	
						End If

					End If
					
					
		
		
					If lgcTB_3_3_3.W4 = "07" Then	 '노무비 =코드 08 + 09 + 10
						oRs2.Find "W4 = '08'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '09'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '10'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2 + dblAmt3 Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "노무비","코드 08 + 09 + 10"))
						   blnError = True	
						End If

					End If
		 
		
		
		
		
		
					If lgcTB_3_3_3.W4 = "12" Then   ' 경비 = 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25+ 26 + 27
						
						oRs2.Find "W4 = '13'"
						dblAmt1 =  UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '14'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '15'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '16'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '17'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '18'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '19'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '20'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '21'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '22'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '23'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '24'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '25'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '26'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '27'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						
						' -- 2006.03 개정 
						oRs2.MoveFirst
						oRs2.Find "W4 = '35'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1  Then
						   Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "경비","코드 13 + 14 + 15 + 16 + 17 + 18 + 19 + 20 + 21 + 22 + 23 + 24 + 25+ 26 + 27"))
						   blnError = True	
						End If	
						
			
					End If
		
		
					If lgcTB_3_3_3.W4 = "28" Then	 '코드(28)당기총공사비용 =코드 01 + 07 + 11 + 12
						oRs2.MoveFirst				' 이전 레코드로 돌아가서 검색하고자할때 MoveFirst해야됨 
						oRs2.Find "W4 = '01'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '07'"
						dblAmt1 = dblAmt1 +  UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '11'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '12'"
						dblAmt1 = dblAmt1 + UNICDbl(oRs2("W5"), 0)
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1  Then
						    Call SaveHTFError(lgsPGM_ID, UNICDbl(lgcTB_3_3_3.W5, 0), UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당기총공사비용","코드  01 + 07 + 11 + 12"))
							blnError = True	
						End If

					End If

				
					If lgcTB_3_3_3.W4 = "31" Then	 '코드(31)합계= 코드 28 + 29 + 30
						oRs2.Find "W4 = '28'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '29'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						
						oRs2.Find "W4 = '30'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)
						
						
						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 + dblAmt2  + dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "코드(31)합계"," 코드 28 + 29 + 30"))
							blnError = True	
						End If
					End If
		
		
					If lgcTB_3_3_3.W4 = "34" Then	 '코드(34)당기완성주택등공사비 = 코드 31 - 32 - 33
						oRs2.Find "W4 = '31'"		' 해당코드는 반드시 존재해야, 다음행에서 에러가 안남 
						dblAmt1 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '32'"
						dblAmt2 = UNICDbl(oRs2("W5"), 0)
						
						oRs2.Find "W4 = '33'"
						dblAmt3 = UNICDbl(oRs2("W5"), 0)

						If UNICDbl(lgcTB_3_3_3.W5, 0) <> dblAmt1 - dblAmt2 - dblAmt3 Then
						    Call SaveHTFError(lgsPGM_ID, lgcTB_3_3_3.W5, UNIGetMesg(TYPE_CHK_NOT_EQUAL, "당기완성주택등공사비"," 코드 31 - 32 - 33"))
							blnError = True	
						End If

					End If
						
		Else
		       blnError = True	
		End if
		
		' -- 2006.03 개정 
		Select Case lgcTB_3_3_3.W4
			Case "35"
				arrNew(35) = lgcTB_3_3_3.W5
			Case Else
				sHTFBody = sHTFBody & UNINumeric(lgcTB_3_3_3.W5, 15, 0)
		End Select

		lgcTB_3_3_3.MoveNext 

	Loop

	' -- 2006.03 개정서식 
	sHTFBody = sHTFBody & UNINumeric(arrNew(35), 15, 0)
	
	
	sHTFBody = sHTFBody & UNIChar("", 19)	' -- 공란 
	
	' ----------- 
	PrintLog "WriteLine2File : " & sHTFBody
	' -- 파일에 기록한다.
	If Not blnError Then
		Call WriteLine2File(sHTFBody)
	End If
	
	Set cDataExists = Nothing	' -- 메모리해제 
	Set lgcTB_3_3_3 = Nothing	' -- 메모리해제 
	
End Function


' ------------------ 조회 함수 --------------------------------
Sub SubMakeSQLStatements_W1113MA1(pMode, pCode1, pCode2, pCode3)
    Select Case pMode 
	  
	  Case "A103" '-- 외부 참조 SQL

	End Select
	PrintLog "SubMakeSQLStatements_W1113MA1 : " & lgStrSQL
End Sub
%>
