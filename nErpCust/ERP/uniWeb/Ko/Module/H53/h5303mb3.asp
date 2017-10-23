<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

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
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	strFileName = lgKeyStream(2) 
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
%>

<script language="vbscript">

'	On Error Resume Next
    Dim lgstrData
    Function FileRead()
		Dim FSO
		Dim FSet
		Dim strLine
		Dim varExist
		Dim res_no
		Dim strSelect,strFrom, strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
		Dim iKey1
		Dim str_data
		Dim cnt
'------------------
'		Dim iColSep, iRowSep
		     
		Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
		Dim iFormLimitByte						'102399byte
		 		
		Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
			
		Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
		Dim iTmpCUBufferCount					'현재의 버퍼 Position
		Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

'		iColSep = parent.gColSep : iRowSep = parent.gRowSep 
		 	
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

		Set FSO = CreateObject("Scripting.FileSystemObject")	

		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		varExist = FSO.FileExists("<%= strFileName %>")
		If varExist = False Then
			Call DisplayMsgBox("115191", "X", "X", "X")
		    Exit Function
		End if

		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		Set FSet = FSO.OpenTextFile("<%= strFileName %>" ,1)			
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		' 여기쯤에서 위에서 열은 txt 파일이 실제 디스켓생성때 만든 txt 파일인지 검사하는 루틴이 
		' 필요할것 같네여. txt 파일 포멧을 검사한다던가..등등...
	'
		iKey1 = FilterVar("<%=lgKeyStream(0)%>", "''", "S")
		
        strSelect = " d.income_tot_amt + d.non_tax5 - ISNULL(SUM(ISNULL(c.a_pay_tot_amt,0)),0) - ISNULL(SUM(ISNULL(c.a_bonus_tot_amt,0)),0) - ISNULL(SUM(ISNULL(c.a_after_bonus_amt,0)),0)   income_tot_amt,  "
        strSelect = strSelect & " CASE WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, a.entr_dt), 112) < " & iKey1 & " + " & FilterVar("0101", "''", "S") & " "
        strSelect = strSelect & " THEN DATEDIFF(month, CONVERT(DATETIME, " & iKey1 & " + " & FilterVar("0101", "''", "S") & "), CONVERT(DATETIME, " & iKey1 & " + " & FilterVar("1231", "''", "S") & ")) + 1 "
        strSelect = strSelect & " WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, a.entr_dt), 112) >= " & iKey1 & " + " & FilterVar("0101", "''", "S") & " "
        strSelect = strSelect & " THEN DATEDIFF(month, ISNULL(b.med_acq_dt, a.entr_dt), CONVERT(DATETIME, " & iKey1 & " + " & FilterVar("1231", "''", "S") & ")) + 1 END  work_month_amt "
        strFrom =  " ((haa010t a left join hdf020t b on  a.emp_no = b.emp_no ) left join hfa050t d on  a.emp_no = d.emp_no) left join hfa040t c  on  a.emp_no = c.emp_no and d.year_yy = c.year_yy"
		
		lgstrData = ""
		str_data = ""
		cnt = 1
		Do While Not FSet.AtEndOfStream
			strLine = ""	
			res_no = ""	
			strWhere = ""	
					
			strLine = FSet.ReadLine
			strLine = rtrim(strLine)
			lgstrData = lgstrData & Chr(11) & mid(strLine,1,6)		'연번 
			lgstrData = lgstrData & Chr(11) & mid(strLine,7,8)		'사업장기호 
			lgstrData = lgstrData & Chr(11) & mid(strLine,15,1)		'차수 
			lgstrData = lgstrData & Chr(11) & mid(strLine,16,2)		'회계 
			lgstrData = lgstrData & Chr(11) & mid(strLine,18,3)		'단위 사업장 
			lgstrData = lgstrData & Chr(11) & mid(strLine,21,11)	'증번호 
'			lgstrData = lgstrData & Chr(11) & mid(strLine,32,len(strLine)-67)	'성명 
'			res_no = left(right(strLine,36),13)			
'			lgstrData = lgstrData & Chr(11) & res_no	'주민번호 
'			lgstrData = lgstrData & Chr(11) & left(right(strLine,23),8)		'자격 취득일 
'			lgstrData = lgstrData & Chr(11) & left(right(strLine,15),2)		'전년도보험료 납부월수 
'			lgstrData = lgstrData & Chr(11) & right(strLine,13)				'전년도보험료 남부총액 
'------------
			lgstrData = lgstrData & Chr(11) & mid(strLine,32,len(strLine)-84)	'성명 
			res_no = left(right(strLine,53),13)			
			lgstrData = lgstrData & Chr(11) & res_no	'주민번호 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,40),8)		'자격 취득일 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,32),2)		'전년도보험료 납부월수 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,30),13)		'전년도보험료 남부총액 
'---------------
			strWhere = " replace(a.res_no,'-','') = " & FilterVar(res_no, "''", "S")
			strWhere = strWhere & " And d.year_yy = " & iKey1
			strWhere = strWhere & " Group By b.med_acq_dt, a.entr_dt,d.income_tot_amt, d.non_tax5 "

			Call parent.CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF0,Chr(11),""))		'전년도보수총액 
			lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF1,Chr(11),""))		'근무월수 

			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			
			str_data	=  strLine 
			str_data	=  str_data & SetFixSrting(Trim(Replace(lgF0,Chr(11),"")),"","0",15,"RIGHT")
			str_data	=  str_data & SetFixSrting(Trim(Replace(lgF1,Chr(11),"")),"","0",2,"RIGHT") & Chr(11)	
			
'-----------------
 			If strCUTotalvalLen + Len(str_data) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
 			   Set objTEXTAREA = parent.document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
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
			cnt = cnt + 1
			
			If Err.Number <> 0 Then
				Call DisplayMsgBox("115100", "X", "X", "X")
				Exit Function
			End If

		Loop
				
'------------------			

 		If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
		   Set objTEXTAREA = parent.document.createElement("TEXTAREA")
 		   objTEXTAREA.name   = "txtCUSpread"
 		   objTEXTAREA.value = Join(iTmpCUBuffer,"")

 		   parent.divTextArea.appendChild(objTEXTAREA)

 		End If
'------------------	 	

		Set FSet = Nothing
		Set FSO =  Nothing
		
		FileRead = True
	End Function

	If Not FileRead() Then
		Call DisplayMsgBox("115100", "X", "X", "X")
	End If
	
  '  Call parent.DbQueryOk_one()

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
    Else                                  '입력값이 존재하면서 한글일경우 
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2                                                  '한글부분만 길이를 각각 2로한다.
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
                .FileOK
	         End with
          End If   
</script>