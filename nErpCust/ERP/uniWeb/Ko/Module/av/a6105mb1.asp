<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<% 
Call LoadBasisGlobalInf() 

On Error Resume Next
Err.Clear

Dim strFileName
Dim strFlag
Dim strMode

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strFileName = Request("hFileName")
strFlag  = Request("cboFlag") 

Call HideStatusWnd

Select Case strMode
    Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

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
		' 여기쯤에서 위에서 열은 txt 파일이 실제 디스켓생성때 만든 txt 파일인지 검사하는 루틴이 
		' 필요할것 같네여. txt 파일 포멧을 검사한다던가..등등...
		'
	    loopCnt  = 0
        parent.strTmpGrid = ""
        parent.strTmpGrid1 = ""
		Do While Not FSet.AtEndOfStream
			strLine = FSet.ReadLine
			If sFlag = "I" Then	'매입 
		        Select Case Mid(strLine, 1, 1)
		            Case 7	'표지 
		                Call parent.subCompany(strLine)
		            Case 2	'매입처정보 
                        loopCnt  = loopCnt + 1
		                Call parent.subRceipt(strLine,loopCnt)
		            Case 4	'매입처합계정보 
		                Call parent.subRceiptSum(strLine)
		        End Select
			ElseIf sFlag = "O" Then	'매출 
		        Select Case Mid(strLine, 1, 1)
		            Case 7	'표지 
		                Call parent.subCompany(strLine)
		            Case 1	'매출처정보 
                        loopCnt  = loopCnt + 1
		                Call parent.subPayment(strLine,loopCnt)
		            Case 3	'매출처합계정보 
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

