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

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	strFileName = lgKeyStream(2) 
    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
		     
		Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
		Dim iFormLimitByte						'102399byte
		 		
		Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
			
		Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
		Dim iTmpCUBufferCount					'������ ���� Position
		Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

'		iColSep = parent.gColSep : iRowSep = parent.gRowSep 
		 	
		'�ѹ��� ������ ������ ũ�� ���� 
		iTmpCUBufferMaxCount = parent.parent.C_CHUNK_ARRAY_COUNT	
		     
		'102399byte
		iFormLimitByte = parent.parent.C_FORM_LIMIT_BYTE
			     
			'������ �ʱ�ȭ 
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

		' �����뿡�� ������ ���� txt ������ ���� ���ϻ����� ���� txt �������� �˻��ϴ� ��ƾ�� 
		' �ʿ��Ұ� ���׿�. txt ���� ������ �˻��Ѵٴ���..���...
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
			lgstrData = lgstrData & Chr(11) & mid(strLine,1,6)		'���� 
			lgstrData = lgstrData & Chr(11) & mid(strLine,7,8)		'������ȣ 
			lgstrData = lgstrData & Chr(11) & mid(strLine,15,1)		'���� 
			lgstrData = lgstrData & Chr(11) & mid(strLine,16,2)		'ȸ�� 
			lgstrData = lgstrData & Chr(11) & mid(strLine,18,3)		'���� ����� 
			lgstrData = lgstrData & Chr(11) & mid(strLine,21,11)	'����ȣ 
'			lgstrData = lgstrData & Chr(11) & mid(strLine,32,len(strLine)-67)	'���� 
'			res_no = left(right(strLine,36),13)			
'			lgstrData = lgstrData & Chr(11) & res_no	'�ֹι�ȣ 
'			lgstrData = lgstrData & Chr(11) & left(right(strLine,23),8)		'�ڰ� ����� 
'			lgstrData = lgstrData & Chr(11) & left(right(strLine,15),2)		'���⵵����� ���ο��� 
'			lgstrData = lgstrData & Chr(11) & right(strLine,13)				'���⵵����� �����Ѿ� 
'------------
			lgstrData = lgstrData & Chr(11) & mid(strLine,32,len(strLine)-84)	'���� 
			res_no = left(right(strLine,53),13)			
			lgstrData = lgstrData & Chr(11) & res_no	'�ֹι�ȣ 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,40),8)		'�ڰ� ����� 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,32),2)		'���⵵����� ���ο��� 
			lgstrData = lgstrData & Chr(11) & left(right(strLine,30),13)		'���⵵����� �����Ѿ� 
'---------------
			strWhere = " replace(a.res_no,'-','') = " & FilterVar(res_no, "''", "S")
			strWhere = strWhere & " And d.year_yy = " & iKey1
			strWhere = strWhere & " Group By b.med_acq_dt, a.entr_dt,d.income_tot_amt, d.non_tax5 "

			Call parent.CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF0,Chr(11),""))		'���⵵�����Ѿ� 
			lgstrData = lgstrData & Chr(11) & Trim(Replace(lgF1,Chr(11),""))		'�ٹ����� 

			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			
			str_data	=  strLine 
			str_data	=  str_data & SetFixSrting(Trim(Replace(lgF0,Chr(11),"")),"","0",15,"RIGHT")
			str_data	=  str_data & SetFixSrting(Trim(Replace(lgF1,Chr(11),"")),"","0",2,"RIGHT") & Chr(11)	
			
'-----------------
 			If strCUTotalvalLen + Len(str_data) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
 			   Set objTEXTAREA = parent.document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
			   objTEXTAREA.name = "txtCUSpread"
 			   objTEXTAREA.value = Join(iTmpCUBuffer,"")
 			   parent.divTextArea.appendChild(objTEXTAREA)     
 			 
 			   iTmpCUBufferMaxCount = parent.parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
			   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
 			   iTmpCUBufferCount = -1
 			   strCUTotalvalLen  = 0
 			End If
			       
 			iTmpCUBufferCount = iTmpCUBufferCount + 1
 			      
 			If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
			   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
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

 		If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
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
' Name : SetFixSrting(�Է°�,�񱳹���,��ü����,��������,�������Ĺ���)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '�Է°��� ��������������� �Է°��� ���̸� 0���� �Ѵ�.
        Cnt = 0     
    Else                                  '�Է°��� �����ϸ鼭 �ѱ��ϰ�� 
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2                                                  '�ѱۺκи� ���̸� ���� 2���Ѵ�.
            Else
                MCnt = MCnt + 1    
            End If
        Next
        Cnt = MCnt
                         
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '�񱳹��ڰ� ������� 
        Else                                                                         '�񱳹��ڰ� �����Ұ�� �񱳹��ڸ� �� �������� �Է°������Ѵ�.
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
    
    If InPos = "" Then                                                              '�������̰� �������� ���� ��� �Է¹��� ���̰� �������̰� �ȴ�.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '���������ϰ��(defalut) �������� ���� ���� ������ ���ڰ� �ԷµǸ� ������ ����(default)�κ��� ��ü���ڷ� ü���.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '���������� 
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

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