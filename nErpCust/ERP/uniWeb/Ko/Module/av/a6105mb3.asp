<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<% 
Call LoadBasisGlobalInf() 

On Error Resume Next
Err.Clear

Dim strFileName
Dim strMode

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strFileName = Request("hFileName")


Call HideStatusWnd

Select Case strMode
    Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

%>

<script language="vbscript">
	'On Error Resume Next
    
    Function FileRead()
		Dim FSO
		Dim FSet
		Dim strLine
		Dim varExist
        Dim loopCnt

		FileRead = False
		

		Set FSO = CreateObject("Scripting.FileSystemObject")	
		If Err.Number <> 0 Then
		    Msgbox Err.Number & " : " & Err.Description
		    Exit Function
		End If

		varExist = FSO.FileExists("<%= strFileName %>")
		If varExist = False Then
		'	Call parent.DisplayMsgbox("115191", "X", "X", "X")
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

	    loopCnt  = 0
        parent.strTmpGrid7 = ""
        parent.strTmpGrid8 = ""
		Do While Not FSet.AtEndOfStream
			strLine = FSet.ReadLine
		        Select Case Mid(strLine, 1, 1)
		            Case "A"   '��� 
		                Call parent.subCompany3(strLine)
		            Case "B"	'�հ����� 
		                Call parent.subExportSum(strLine)
		            Case "C"	'����������� 
                        loopCnt  = loopCnt + 1
		                Call parent.subExportList(strLine,loopCnt)
		        End Select
			If Err.Number <> 0 Then
				Call parent.DisplayMsgbox("800186", "X", "X", "X")
				Exit Function
			End If
			
		Loop

		Set FSet = Nothing
		Set FSO = Nothing
		
		FileRead = True
	End Function
	
	If Not FileRead() Then
		Call parent.DisplayMsgbox("800186", "X", "X", "X")
	End If
    Call parent.DbQueryOk_three()

</script>

<%
End Select
Response.End
%>
