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

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	strFileName		  = lgKeyStream(3) 
	strFileGubun	  = Request("htxtFileGubun")  	
'    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
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
		     
		Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
		Dim iFormLimitByte						'102399byte
		 		
		Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ����
			
		Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
		Dim iTmpCUBufferCount					'������ ���� Position
		Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

		
		iColSep = parent.parent.gColSep : iRowSep = parent.parent.gRowSep 
		 	
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

		Set ws = wb.Worksheets(1) 'Worksheet ��ü ����
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If		
	
		Set objRange = ws.UsedRange '���� ���� ��ü ����
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If	
	
		aData= objRange.value '���� ������ ������ 2�����迭 aData�� �ѱ�
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

		For i=2 to  uBound(aData, 1) ' �迭�� ������ �����

			If Strprovcd = "%" Then
				Select Case "<%= strFileGubun %>"
					   Case "A"
					   		If (StrDt = Trim(aData(i,6))) And (StrYYMM = Trim(aData(i,4))) Then
								lgstrData = lgstrData & "C" & Chr(11) & i-1
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�μ��ڵ�		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���	
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�ش���
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,6)	'������
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'�޿�
								lgstrData = lgstrData & Chr(11) & aData(i,8)	'��
								lgstrData = lgstrData & Chr(11) & aData(i,9)	'������Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,10)	'�����Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,11)	'�����Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,12)	'��������
								lgstrData = lgstrData & Chr(11) & aData(i,13)	'�����޾�
								lgstrData = lgstrData & Chr(11) & aData(i,14)	'�ҵ漼
								lgstrData = lgstrData & Chr(11) & aData(i,15)	'�ֹμ�
								lgstrData = lgstrData & Chr(11) & aData(i,16)	'���ο���
								lgstrData = lgstrData & Chr(11) & aData(i,17)	'�ǰ�����
								lgstrData = lgstrData & Chr(11) & aData(i,18)	'��뺸��
								lgstrData = lgstrData & Chr(11) & Chr(12)
							End If							
	
					   Case "B"
					   		If (StrYYMM = Trim(aData(i,2))) Then				   
								lgstrData = lgstrData & "C" & Chr(11) & i-1							
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�ش���		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�����ڵ�
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'����ݾ�
								lgstrData = lgstrData & Chr(12)
							End If
							
					   Case "C"
					   		If (StrYYMM = Trim(aData(i,2))) Then					   		
								lgstrData = lgstrData & "C" & Chr(11) & i-1						
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�ش���		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�����ڵ�
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'�����ݾ�
								lgstrData = lgstrData & Chr(12)
							End If
				End Select
			Else
				Select Case "<%= strFileGubun %>"
					   Case "A"
					   		If (StrDt = Trim(aData(i,6))) And (StrYYMM = Trim(aData(i,4))) And (Strprovcd = Trim(aData(i,5))) Then
								lgstrData = lgstrData & "C" & Chr(11) & i-1
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�μ��ڵ�		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���	
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�ش���
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,6)	'������
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'�޿�
								lgstrData = lgstrData & Chr(11) & aData(i,8)	'��
								lgstrData = lgstrData & Chr(11) & aData(i,9)	'������Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,10)	'�����Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,11)	'�����Ѿ�
								lgstrData = lgstrData & Chr(11) & aData(i,12)	'��������
								lgstrData = lgstrData & Chr(11) & aData(i,13)	'�����޾�
								lgstrData = lgstrData & Chr(11) & aData(i,14)	'�ҵ漼
								lgstrData = lgstrData & Chr(11) & aData(i,15)	'�ֹμ�
								lgstrData = lgstrData & Chr(11) & aData(i,16)	'���ο���
								lgstrData = lgstrData & Chr(11) & aData(i,17)	'�ǰ�����
								lgstrData = lgstrData & Chr(11) & aData(i,18)	'��뺸��
								lgstrData = lgstrData & Chr(11) & Chr(12)
							End If							
	
					   Case "B"
					   		If (StrYYMM = Trim(aData(i,2))) And (Strprovcd = Trim(aData(i,4))) Then	
								lgstrData = lgstrData & "C" & Chr(11) & i-1							
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�ش���		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�����ڵ�
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'����ݾ�
								lgstrData = lgstrData & Chr(12)
							End If
							
					   Case "C"
					   		If (StrYYMM = Trim(aData(i,2))) And (Strprovcd = Trim(aData(i,4))) Then					   		
								lgstrData = lgstrData & "C" & Chr(11) & i-1						
								lgstrData = lgstrData & Chr(11) & aData(i,2)	'�ش���		
								lgstrData = lgstrData & Chr(11) & aData(i,3)	'���							
								lgstrData = lgstrData & Chr(11) & aData(i,4)	'�޿������ڵ�							
								lgstrData = lgstrData & Chr(11) & aData(i,5)	'�����ڵ�
								lgstrData = lgstrData & Chr(11) & aData(i,7)	'�����ݾ�
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
' Name : SetFixSrting(�Է°�,�񱳹���,��ü����,��������,�������Ĺ���)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '�Է°��� ��������������� �Է°��� ���̸� 0���� �Ѵ�.
        Cnt = 0     
    Else																			'�Է°��� �����ϸ鼭 �ѱ��ϰ��
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2														'�ѱۺκи� ���̸� ���� 2���Ѵ�.
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
