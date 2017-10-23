Sub SaveValueForm()
	Dim goURL
	Dim nCnt 
	Dim i
	Dim	nWindowHeight
	Dim objDoc, objEl
	Dim objCond, objComDtl
	Dim nIsDiv
	
	goURL = ""
	nCnt = 0
	nIsDiv = 0
	
	on error resume next
	Set objDoc = frBody.document

	i = objDoc.All.MyTab.Length

	If i <= 1 Then ' Tab이 하나인 경우 
		For i = 0 To objDoc.All.Length - 1
			If UCase(objDoc.All(i).TagName) = "TD" Then
				If Left(UCase(objDoc.All(i).className), 3) = "TAB" Then
					Set ObjEl = objDoc.All(i)
					Exit For
				End If
			End If
		Next

	Else ' Tab이 하나 이상인 경우 
		For i = 0 To objDoc.All.MyTab.Length - 1
			If objDoc.All.TabDiv(i).Style.display = "" Then
				Set objEl = objDoc.All.TabDiv(i)
				Exit For
			End If
		Next

		' 공통 조건은 반드시 filedset에 쌓여져 있어야 한다. DIV가 나오기 전에 
		For i = 0 To objDoc.All.Length - 1
			If UCase(objDoc.All(i).TagName) = "DIV" Then
				Exit For
			End If

			If UCase(objDoc.All(i).TagName) = "FIELDSET" Then
				Set objCond = objDoc.All(i)
				Exit For
			End If

		Next

		' 공통 싱글 detail a3111ma1 땜시 
		For i = 0 To objDoc.All.Length - 1
			If UCase(objDoc.All(i).TagName) = "DIV" Then
				Set objComDtl = Nothing
				i = i + objDoc.All(i).All.Length
				nIsDiv = -999
			End If
			
			If UCase(objDoc.All(i).TagName) = "TABLE" And nIsDiv = -999 Then
				Set objComDtl = objDoc.All(i)
				Exit For
			End If
		Next

	End If

	SearchText objCond, goURL, nCnt	
	SearchText objEl, goURL, nCnt	
	SearchText objComDtl, goURL, nCnt	
	
	Set objCond = Nothing
	Set objEl = Nothing
	Set objComDtl = Nothing
	Set objDoc = Nothing

	If nCnt = 0 Then
       MsgBox "입력된 값이 없습니다.", vbInformation
       Exit Sub
	End If

	nWindowHeight = (nCnt * 24 + 70) & "px"
	goURL = "ComAsp/DefaultSaveForm.asp?cnt=" & nCnt & goURL
	Call window.showModalDialog(goURL, null, "dialogWidth=230px; dialogHeight=" & nWindowHeight & "; center: Yes; help: No; resizable: Yes; status: No; scrollbars: no")

End Sub

Sub	SearchText(Byval objEl, ByRef goURL, ByRef nCnt)	
	Dim i

	For i = 0 To objEl.All.Length - 1
		Select Case UCase(objEl.All(i).TagName)
			Case "INPUT"
				If UCase(objEl.All(i).Type) = "TEXT" And Trim(objEl.All(i).Value) <> "" Then
					nCnt = nCnt + 1
					If UCase(objEl.All(i).Style.textTransform) = "UPPERCASE" Then
						goURL = goURL & "&txt" & nCnt & "=" & UCase(objEl.All(i).Value)
					ElseIf UCase(objEl.All(i).Style.textTransform) = "LOWERCASE" Then
						goURL = goURL & "&txt" & nCnt & "=" & LCase(objEl.All(i).Value)
					Else
						goURL = goURL & "&txt" & nCnt & "=" & objEl.All(i).Value
					End If
				End If
		End Select
	Next
End Sub  

Sub LoadValueForm()
	Dim nCnt
	Dim nWindowHeight

	nCnt = IsExistCookie("DEFAULT_VALUE")

	If nCnt = -999 Then
	   Msgbox "저장된 값이 없습니다.", vbInformation
	   Exit Sub
	ElseIf nCnt > 20 Then
		nCnt = 20	
	End If

	nWindowHeight = (nCnt * 24 + 70) & "px"

	Call window.showModalDialog("ComAsp/DefaultLoadForm.asp", null, "dialogWidth=230px; dialogHeight=" & nWindowHeight & "; center: Yes; help: No; resizable: Yes; status: No; scrollbars: no")

End Sub

Function IsExistCookie(byval key)
	Dim s, e, myCookie
	Dim nCnt

	IsExistCookie = -999

	myCookie = unescape(Document.Cookie)

	s = Instr(1, myCookie, key & "=")
	If s = 0 Then
		Exit Function
	End If

	s = s + Len(key & "=")

	i = s + 1
	nCnt = 0
    Do While InStr(i, myCookie, "=", vbTextCompare) <> 0
		s = InStr(i, myCookie, "=", vbTextCompare)
		e = InStr(i, myCookie, ";", vbTextCompare)
		If s > e And e <> 0 Then
			Exit Do
		End If

		nCnt = nCnt + 1
        i = InStr(i, myCookie, "=", vbTextCompare) + 1
    Loop

	IsExistCookie = nCnt - 1

End Function

Function get_sequence_checkbox()
	Dim s

	s = ""
	For i = 0 To document.forms(0).cnt.value - 1
		If document.forms(0).elements(element_offset+i).checked = True Then
			If Len(s) > 0 Then
				s = s + "," 
			End If
			s = s + document.forms(0).elements(element_offset+i).value
		End If
	Next

	get_sequence_checkbox = s

End Function
