<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<% 
Call LoadBasisGlobalInf() 

On Error Resume Next
Err.Clear

Dim strFileName
Dim strFlag
Dim strMode

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strFileName = Request("hFileName")
strFlag  = Request("cboFlag") 

Call HideStatusWnd

Select Case strMode
    Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

%>

<script language="vbscript">
	On Error Resume Next
    
    Function FileRead()
		Dim FSO
		Dim FSet
		Dim strLine
		Dim varExist
		Dim sFlag
		Dim loopCnt,loopCnt1,loopCnt2
		
		
		FileRead = False
		sFlag = "<%= strFlag %>"

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

		'
		' shj
		'
		' �����뿡�� ������ ���� txt ������ ���� ���ϻ����� ���� txt �������� �˻��ϴ� ��ƾ�� 
		' �ʿ��Ұ� ���׿�. txt ���� ������ �˻��Ѵٴ���..���...
		'
            parent.strTmpGrid2 = ""
            parent.strTmpGrid3 = ""
            parent.strTmpGrid4 = ""
			loopCnt     = 0
			loopCnt1    = 0
			loopCnt2    = 0
			Do While Not FSet.AtEndOfStream
				strLine = FSet.ReadLine
			    If sFlag = "I" Then	'���� 
			       If Mid(strLine, 1, 1) = "A" Then
							Call parent.subCompany2A(strLine) '//���������A
				   ElseIf  Mid(strLine, 1, 1) = "B"  Then
							loopCnt = loopCnt + 1
							Call parent.subCompany2B(strLine,loopCnt) '//���������B
				   Else
                            	loopCnt1 = loopCnt1 + 1
						Select Case Mid(strLine, 1, 3)
							Case "C18"	'����ó�հ����� 
							    Call parent.subRceiptSum2(strLine,loopCnt)
							Case "D18"	'����ó���� 
                            	loopCnt2 = loopCnt2 + 1
							    Call parent.subRceipt2(strLine, loopCnt)
						End Select
				   End If		
			       
			   	ElseIf sFlag = "O" Then	'���� 
			   		If Mid(strLine, 1, 1) = "A" Then
						Call parent.subCompany2A(strLine) '//���������A
					ElseIf  Mid(strLine, 1, 1) = "B"  Then
						loopCnt = loopCnt + 1
						Call parent.subCompany2B(strLine, loopCnt) '//���������B
					Else
						Select Case Mid(strLine, 1, 3)
						    Case "C17"	'����ó�հ����� 
                            	loopCnt1 = loopCnt1 + 1
						        Call parent.subPaymentSum2(strLine, loopCnt)
						    Case "D17"	'����ó���� 
                            	loopCnt2 = loopCnt2 + 1
						        Call parent.subPayment2(strLine, loopCnt)
						End Select
					End If
				End If		'// input- output
				
				If Err.Number <> 0 Then
					Call DisplayMsgBox("115100", "X", "X", "X")
					Exit Function
				End If
			Loop
		
		Set FSet = Nothing
		Set FSO = Nothing
		
		FileRead = True
	End Function
	
	If Not FileRead() Then
		Call DisplayMsgBox("115100", "X", "X", "X")
	Else
			Call parent.DbQueryOk_two()	
	End If

</script>

<%
End Select
Response.End
%>

