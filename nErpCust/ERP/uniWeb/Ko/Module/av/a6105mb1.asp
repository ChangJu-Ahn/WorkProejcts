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
        Dim loopCnt

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
	    loopCnt  = 0
        parent.strTmpGrid = ""
        parent.strTmpGrid1 = ""
		Do While Not FSet.AtEndOfStream
			strLine = FSet.ReadLine
			If sFlag = "I" Then	'���� 
		        Select Case Mid(strLine, 1, 1)
		            Case 7	'ǥ�� 
		                Call parent.subCompany(strLine)
		            Case 2	'����ó���� 
                        loopCnt  = loopCnt + 1
		                Call parent.subRceipt(strLine,loopCnt)
		            Case 4	'����ó�հ����� 
		                Call parent.subRceiptSum(strLine)
		        End Select
			ElseIf sFlag = "O" Then	'���� 
		        Select Case Mid(strLine, 1, 1)
		            Case 7	'ǥ�� 
		                Call parent.subCompany(strLine)
		            Case 1	'����ó���� 
                        loopCnt  = loopCnt + 1
		                Call parent.subPayment(strLine,loopCnt)
		            Case 3	'����ó�հ����� 
		                Call parent.subPaymentSum(strLine)
		        End Select
			End If
			
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
	End If
    Call parent.DbQueryOk_one()
</script>

<%
End Select
Response.End
%>

