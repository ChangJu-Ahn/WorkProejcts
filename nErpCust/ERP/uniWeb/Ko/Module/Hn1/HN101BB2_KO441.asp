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
		
    Dim StrDt, StrYYMM, StrProvCD, StrFileGubun	
    
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	strFileName		  = lgKeyStream(3) 
	strFileGubun	  = Request("htxtFileGubun")  	
'    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
	StrDt		= Trim(Request("htxtDt"))
	StrYYMM		= Replace(Trim(Request("htxtYYMM")),"-","")
	StrProvCD	= Trim(Request("htxtProvCD"))
	If IsNull(StrProvCD) Or StrProvCD = "" Then StrProvCD = "%"

%>

<script language="vbscript">

	'On Error Resume Next
    Dim lgstrData, lgLngMaxRow
   	Dim StrDt, StrYYMM, StrProvCD
	
	StrDt		= "<%=StrDt		%>"
	StrYYMM		= "<%=StrYYMM	%>"
	Strprovcd	= "<%=StrProvCD	%>"
    
    Function FileRead()

		Dim FSO, wb, ws, objRange
		Dim FSet, aData
		Dim strLine
		Dim varExist
		Dim res_no

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

		If trim(aData(2,1)) = "" Then
			Call DisplayMsgBox("171910", "X", "X", "X")
			Exit Function
		End If

		
		lgstrData = ""
		
		lgLngMaxRow = uBound(aData, 1) - 1

		For i=2 to  uBound(aData, 1) ' 배열의 끝까지 행루프

			If Strprovcd = "%" Then
				Select Case "<%= strFileGubun %>"
					   Case "A"
					   		If (StrDt = Trim(aData(i,6))) And (StrYYMM = Trim(aData(i,4))) Then
								lgstrData = lgstrData & "C" & Chr(11) & i-1
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'부서코드		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번	
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'해당년월
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,6)	'지급일
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'급여
								lgstrData = lgstrData & Chr(11) & aData(i,8)	'상여
								lgstrData = lgstrData & Chr(11) & aData(i,9)	'비과세총액
								lgstrData = lgstrData & Chr(11) & aData(i,10)	'과세총액
								lgstrData = lgstrData & Chr(11) & aData(i,11)	'지급총액
								lgstrData = lgstrData & Chr(11) & aData(i,12)	'제공제계
								lgstrData = lgstrData & Chr(11) & aData(i,13)	'실지급액
								lgstrData = lgstrData & Chr(11) & aData(i,14)	'소득세
								lgstrData = lgstrData & Chr(11) & aData(i,15)	'주민세
								lgstrData = lgstrData & Chr(11) & aData(i,16)	'국민연금
								lgstrData = lgstrData & Chr(11) & aData(i,17)	'건강보험
								lgstrData = lgstrData & Chr(11) & aData(i,18)	'고용보험
								lgstrData = lgstrData & Chr(11) & Chr(12)
							End If							
	
					   Case "B"
					   		If (StrYYMM = Trim(aData(i,2))) Then				   
								lgstrData = lgstrData & "C" & Chr(11) & i-1							
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'수당코드
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'수당금액
								lgstrData = lgstrData & Chr(12)
							End If
							
					   Case "C"
					   		If (StrYYMM = Trim(aData(i,2))) Then					   		
								lgstrData = lgstrData & "C" & Chr(11) & i-1						
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'공제코드
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'공제금액
								lgstrData = lgstrData & Chr(12)
							End If
				End Select
			Else
				Select Case "<%= strFileGubun %>"
					   Case "A"
					   		If (StrDt = Trim(aData(i,6))) And (StrYYMM = Trim(aData(i,4))) And (Strprovcd = Trim(aData(i,5))) Then
								lgstrData = lgstrData & "C" & Chr(11) & i-1
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'부서코드		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번	
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'해당년월
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,6)	'지급일
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'급여
								lgstrData = lgstrData & Chr(11) & aData(i,8)	'상여
								lgstrData = lgstrData & Chr(11) & aData(i,9)	'비과세총액
								lgstrData = lgstrData & Chr(11) & aData(i,10)	'과세총액
								lgstrData = lgstrData & Chr(11) & aData(i,11)	'지급총액
								lgstrData = lgstrData & Chr(11) & aData(i,12)	'제공제계
								lgstrData = lgstrData & Chr(11) & aData(i,13)	'실지급액
								lgstrData = lgstrData & Chr(11) & aData(i,14)	'소득세
								lgstrData = lgstrData & Chr(11) & aData(i,15)	'주민세
								lgstrData = lgstrData & Chr(11) & aData(i,16)	'국민연금
								lgstrData = lgstrData & Chr(11) & aData(i,17)	'건강보험
								lgstrData = lgstrData & Chr(11) & aData(i,18)	'고용보험
								lgstrData = lgstrData & Chr(11) & Chr(12)
							End If							
	
					   Case "B"
					   		If (StrYYMM = Trim(aData(i,2))) And (Strprovcd = Trim(aData(i,4))) Then	
								lgstrData = lgstrData & "C" & Chr(11) & i-1							
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'수당코드
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'수당금액
								lgstrData = lgstrData & Chr(12)
							End If
							
					   Case "C"
					   		If (StrYYMM = Trim(aData(i,2))) And (Strprovcd = Trim(aData(i,4))) Then					   		
								lgstrData = lgstrData & "C" & Chr(11) & i-1						
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'해당년월		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'사번							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'급여유형코드							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'공제코드
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'공제금액
								lgstrData = lgstrData & Chr(12)
							End If
				End Select								
			End If

'------------------ 
			
			If Err.Number <> 0 Then
				Call DisplayMsgBox("115100", "X", "X", "X")
				Exit Function
			End If

		Next
		
	
		FSO.quit
		Set objRange =  Nothing
		Set ws =  Nothing
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
                '.ggoSpread.Source     = .frm1.vspdData
                '.lgStrPrevKey    = "<%=lgStrPrevKey%>"	
'MsgBox "lgLngMaxRow = " & CStr(lgLngMaxRow)
'MsgBox .frm1.txtMaxRows.value               			
                .frm1.txtSpread.value = lgstrData
                .frm1.txtMaxRows.value = lgLngMaxRow         
                .DbSave
                '.ggoSpread.SSShowData lgstrData				
                '.DBAutoQueryOk        
	         End with
          End If   
</script>
