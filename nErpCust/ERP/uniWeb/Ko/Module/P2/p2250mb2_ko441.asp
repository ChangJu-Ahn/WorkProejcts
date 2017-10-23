<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<%
	Dim  lgStrPrevKey
	Dim	 strFileName
	Dim  strFileGubun
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	strFileName		  = lgKeyStream(3) 
	strFileGubun	  = Request("htxtFileGubun")  	
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

'Call ServerMesgBox("strFileName : " & strFileName , vbInformation, I_MKSCRIPT)
		
	'Call ServerMesgBox(strFileGubun , vbInformation, I_MKSCRIPT)
%>

<script language="vbscript">

	'On Error Resume Next
    Dim lgstrData
    Dim lgstrData_header
    
    Function FileRead()
		Dim FSO, wb, ws, objRange
		Dim FSet, aData
		Dim strLine
		Dim varExist
		Dim res_no
		Dim strSelect,strFrom, strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
		Dim iKey1
		Dim str_data
		Dim cnt
'------------------
		Dim iColSep, iRowSep
		     
		Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
		Dim iFormLimitByte						'102399byte
		 		
		Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼
			
		Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
		Dim iTmpCUBufferCount					'현재의 버퍼 Position
		Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

		
		iColSep = parent.parent.gColSep : iRowSep = parent.parent.gRowSep 
		 	
		'한번에 설정한 버퍼의 크기 설정
		iTmpCUBufferMaxCount = parent.parent.C_CHUNK_ARRAY_COUNT	
		     
		'102399byte
		iFormLimitByte = parent.parent.C_FORM_LIMIT_BYTE
		     
		'버퍼의 초기화
		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			

		iTmpCUBufferCount = -1  

		strCUTotalvalLen = 0 
'------------------		
		
		FileRead = False
		
		Set FSO = CreateObject("Excel.Application")
		 
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		Set wb = FSO.Workbooks.Open("<%= strFileName %>")

		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		Set ws = wb.Worksheets(1) 'Worksheet 객체 생성
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If		
	
		Set objRange = ws.UsedRange '사용된 영역 객체 생성
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If	
	
		aData= objRange.value '사용된 영역의 값들을 2차원배열 aData로 넘김
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		If trim(aData(5,3)) = "" Then
			Call DisplayMsgBox("171910", "X", "X", "X")
			Exit Function
		End If

		'iKey1 = FilterVar(lgKeyStream(0), "''", "S")
		
		lgstrData = ""
		lgstrData_header = ""   '20080302::hanc
		cnt = 0


		For i=5 to  uBound(aData, 1) ' 배열의 끝까지 행루프
			str_data = ""	
			strWhere = ""				
		
			Select Case "<%= strFileGubun %>"
				   Case "A"
                        
                        if trim(aData(i,3))  <> "" then

    						lgstrData = lgstrData & Chr(11) & aData(i,2)	        'custom		
    						lgstrData = lgstrData & Chr(11) & ""
    						lgstrData = lgstrData & Chr(11) & aData(i,3)	        '품목코드		
    						lgstrData = lgstrData & Chr(11) & aData(i,4)	        '품목명
    						lgstrData = lgstrData & Chr(11) & "*"	                'Tracking No
    						lgstrData = lgstrData & Chr(11) & "수주계획(수주일)"	'수주계획(수주일)		
    						lgstrData = lgstrData & Chr(11) & "수주계획(수주일)"	'수주계획(수주일)		
    						lgstrData = lgstrData & Chr(11) & aData(i,5)	'1	
    						lgstrData = lgstrData & Chr(11) & aData(i,6)	'2
    						lgstrData = lgstrData & Chr(11) & aData(i,7)	'3
    						lgstrData = lgstrData & Chr(11) & aData(i,8)	'4
    						lgstrData = lgstrData & Chr(11) & aData(i,9)	'5
    						lgstrData = lgstrData & Chr(11) & aData(i,10)	'6
    						lgstrData = lgstrData & Chr(11) & aData(i,11)	'7
    						lgstrData = lgstrData & Chr(11) & aData(i,12)	'8
    						lgstrData = lgstrData & Chr(11) & aData(i,13)	'9
    						lgstrData = lgstrData & Chr(11) & aData(i,14)	'10
    						lgstrData = lgstrData & Chr(11) & aData(i,15)	'11
    						lgstrData = lgstrData & Chr(11) & aData(i,16)	'12
    						lgstrData = lgstrData & Chr(11) & aData(i,17)	'13
    						lgstrData = lgstrData & Chr(11) & aData(i,18)	'14
    						lgstrData = lgstrData & Chr(11) & aData(i,19)	'15
    						lgstrData = lgstrData & Chr(11) & aData(i,20)	'16
    						lgstrData = lgstrData & Chr(11) & aData(i,21)	'17
    						lgstrData = lgstrData & Chr(11) & aData(i,22)	'18
    						lgstrData = lgstrData & Chr(11) & aData(i,23)	'19
    						lgstrData = lgstrData & Chr(11) & aData(i,24)	'20
    						lgstrData = lgstrData & Chr(11) & aData(i,25)	'21
    						lgstrData = lgstrData & Chr(11) & aData(i,26)	'22
    						lgstrData = lgstrData & Chr(11) & aData(i,27)	'23
    						lgstrData = lgstrData & Chr(11) & aData(i,28)	'24
    						lgstrData = lgstrData & Chr(11) & aData(i,29)	'25
    						lgstrData = lgstrData & Chr(11) & aData(i,30)	'26
    						lgstrData = lgstrData & Chr(11) & aData(i,31)	'27
    						lgstrData = lgstrData & Chr(11) & aData(i,32)	'28
    						lgstrData = lgstrData & Chr(11) & aData(i,33)	'29
    						lgstrData = lgstrData & Chr(11) & aData(i,34)	'30
    
    						lgstrData = lgstrData & Chr(11) & aData(i,35)	'31
    						lgstrData = lgstrData & Chr(11) & aData(i,36)	'32
    						lgstrData = lgstrData & Chr(11) & aData(i,37)	'33
    						lgstrData = lgstrData & Chr(11) & aData(i,38)	'34
    						lgstrData = lgstrData & Chr(11) & aData(i,39)	'35
    						lgstrData = lgstrData & Chr(11) & aData(i,40)	'36
    						lgstrData = lgstrData & Chr(11) & aData(i,41)	'37
    						lgstrData = lgstrData & Chr(11) & aData(i,42)	'38
    						lgstrData = lgstrData & Chr(11) & aData(i,43)	'39
    						lgstrData = lgstrData & Chr(11) & aData(i,44)	'30
    
    						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
    						lgstrData = lgstrData & Chr(11) & Chr(12)
    						
    						cnt = cnt + 1
    						
    						str_data = aData(i,2) & iColSep & aData(i,3) & iColSep & aData(i,4) & iColSep & lgF0 & iColSep & aData(i,5) & iColSep & aData(i,6) & iColSep & aData(i,7) & iColSep & aData(i,8) & iColSep  & aData(i,10) & iColSep & aData(i,11) & iColSep & aData(i,12) & iColSep & aData(i,13) & iColSep & aData(i,14) &  iColSep & aData(i,15) & iColSep & aData(i,16) & aData(i,17) & iColSep & aData(i,18) & iColSep & aData(i,19) & iRowSep
                        end if
                        					

				   Case "B"
						strWhere = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(aData(i,4) , "''", "S")
						Call parent.CommonQueryRs("MINOR_NM", "B_MINOR",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						
						lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
						lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
						lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF0,Chr(11),""))	'급여유형코드		
						lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
						lgstrData = lgstrData & Chr(11) & aData(i,5)	'수당코드
						lgstrData = lgstrData & Chr(11) & aData(i,6)	'수당명
						lgstrData = lgstrData & Chr(11) & aData(i,7)	'수당금앨
						
						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)
						
						cnt = cnt + 1
						
						str_data = aData(i,2) & iColSep & aData(i,3) & iColSep & lgF0 & iColSep & aData(i,4) & iColSep & aData(i,5) & iColSep & aData(i,6) & iColSep & aData(i,7) & iRowSep 

				   Case "C"
						strWhere = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(aData(i,4) , "''", "S")
						Call parent.CommonQueryRs("MINOR_NM", "B_MINOR",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						
						lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
						lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
						lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF0,Chr(11),""))	'급여유형코드		
						lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
						lgstrData = lgstrData & Chr(11) & aData(i,5)	'공제코드
						lgstrData = lgstrData & Chr(11) & aData(i,6)	'공제명
						lgstrData = lgstrData & Chr(11) & aData(i,7)	'공제금액
						
						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)
						
						cnt = cnt + 1
						
						str_data = aData(i,2) & iColSep & aData(i,3) & iColSep & lgF0 & iColSep & aData(i,4) & iColSep & aData(i,5) & iColSep & aData(i,6) & iColSep & aData(i,7) & iRowSep 

			End Select
			

 			If strCUTotalvalLen + Len(str_data) >  iFormLimitByte Then			'한개의 form element에 넣을 Data 한개치가 넘으면
			                            
 			   Set objTEXTAREA = parent.document.createElement("TEXTAREA")      '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음
			   objTEXTAREA.name = "txtCUSpread"
 			   objTEXTAREA.value = Join(iTmpCUBuffer,"")
 			   parent.divTextArea.appendChild(objTEXTAREA)     
 			 
 			   iTmpCUBufferMaxCount = parent.parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화
			   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
 			   iTmpCUBufferCount = -1
 			   strCUTotalvalLen  = 0
 			End If
			       
 			iTmpCUBufferCount = iTmpCUBufferCount + 1
 			      
 			If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면
			   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성
			   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
 			End If   
 			         
 			iTmpCUBuffer(iTmpCUBufferCount) =  str_data         
 			strCUTotalvalLen = strCUTotalvalLen + Len(str_data)

'------------------ 
			
			If Err.Number <> 0 Then
				Call DisplayMsgBox("115100", "X", "X", "X")
				Exit Function
			End If
		Next



'20080302::hanc
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,5)	        'day1
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,6)	        'day2
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,7)	        'day3
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,8)	        'day4
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,9)	        'day5
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,10)	        'day6
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,11)	        'day7
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,12)	        'day8
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,13)	        'day9
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,14)	        'day10
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,15)	        'day11
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,16)	        'day12
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,17)	        'day13
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,18)	        'day14
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,19)	        'day15
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,20)	        'day16
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,21)	        'day17
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,22)	        'day18
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,23)	        'day19
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,24)	        'day20
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,25)	        'day21
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,26)	        'day22
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,27)	        'day23
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,28)	        'day24
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,29)	        'day25
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,30)	        'day26
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,31)	        'day27
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,32)	        'day28
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,33)	        'day29
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,34)	        'day30
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,35)	        'day31
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,36)	        'day32
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,37)	        'day33
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,38)	        'day34
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,39)	        'day35
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,40)	        'day36
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,41)	        'day37
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,42)	        'day38
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,43)	        'day39
		lgstrData_header = lgstrData_header & Chr(11) & aData(4,44)	        'day40
		lgstrData_header = lgstrData_header & Chr(11) & Chr(12)

		
'------------------			
 		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리
		   Set objTEXTAREA = parent.document.createElement("TEXTAREA")
 		   objTEXTAREA.name   = "txtCUSpread"
 		   objTEXTAREA.value = Join(iTmpCUBuffer,"")

 		   parent.divTextArea.appendChild(objTEXTAREA)

 		End If
'------------------	 	
		Set objRange =  Nothing
		Set ws =  Nothing
        wb.Close
		Set wb =  Nothing
		Set FSO = Nothing
		
		FileRead = True		
					
	End Function
	
	If Not FileRead() Then
		Call DisplayMsgBox("115100", "X", "X", "X")
	End If


'============================================================================================================
' Name : SetFixSrting(입력값,비교문자,대체문자,고정길이,문자정렬방향)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '입력값이 존재하지않을경우 입력값의 길이를 0으로 한다.
        Cnt = 0     
    Else																			'입력값이 존재하면서 한글일경우
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2														'한글부분만 길이를 각각 2로한다.
            Else
                MCnt = MCnt + 1    
            End If
        Next
        Cnt = MCnt
                         
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '비교문자가 없을경우
        Else                                                                         '비교문자가 존재할경우 비교문자를 뺀 나머지를 입력값으로한다.
            ExSymbol = Split(InValue,ComSymbol)
            If UBound(ExSymbol) > 0 Then
                iDx = UBound(ExSymbol)
                For i = 0 To iDx
                    strSplit = strSplit & ExSymbol(i)
                Next
                InValue = strSplit
            End If               
        End If        
    End If        
    
    If InPos = "" Then                                                              '고정길이가 정해지지 않을 경우 입력문자 길이가 고정길이가 된다.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '왼쪽정렬일경우(defalut) 고정길이 보다 작은 길이의 문자가 입력되면 나머지 공백(default)부분을 대체문자로 체운다.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '오른쪽정렬
        If InPos > Cnt Then                                                           ' ex:     hi
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = 1 To (InPos - Cnt)
                strTemp = strTemp & strFix         
            Next
            InValue = strTemp & InValue
        End If
    End If
    SetFixSrting = InValue
End Function

</script>
<%

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<script language="vbscript">
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"				
                .ggoSpread.SSShowData lgstrData
                .SetHeader(lgstrData_header)            '20080302::hanc    
                .DBAutoQueryOk        
	         End with
          End If   
</script>
