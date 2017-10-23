Dim gblnWinEvent

Function FncExport(Byval iTypeForm)
	Dim SF

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	Call frbody.LayerShowHide(1)

	Set SF = CreateObject("uni2kCM.SaveFile")
	
	Call SF.SaveFile(frBody.document)
	Set SF = Nothing

	Call frbody.LayerShowHide(0)
	gblnWinEvent = False
End Function

Function FncPrint()
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	window.showModalDialog "ComASP/Print.asp", frBody.document, _
		"dialogWidth=" & window.screen.width - 100 & "px; dialogHeight=" & window.screen.height-120 & "px; center=yes; help: No; resizable:Yes; status:No;"
	
	gblnWinEvent = False
End Function
