Const PGM_FIND_SCREEN_ASP = "ComASP/Find.asp"		' Window Screen ASP Name
Dim gobjWin		' Variabe that contains information about Previous Window because of Modaless Window
Dim gFormType	' Type Definition of developer Window
Dim gblnTab		' if Tab of developer Window exists or not

Function FncFind(Byval iType, Byval blnTab)
	Dim x, y
	Dim arrDoc(0)
	Dim arrFormType(0)
	Dim arrInTab(0)
	
	gFormType = iType
	gblnTab = blnTab
	
	x = ( screen.width - 380 ) / 2
	y = ( screen.height - 150 ) / 2
		
	Set gobjWin = Nothing	

    Set arrDoc(0) = frBody
	arrFormType(0) = gFormType	' Window Type
	arrInTab(0) = gblnTab		'  if Tab exists or not
	gobjWin = window.showModalDialog(PGM_FIND_SCREEN_ASP, Array(arrDoc, arrFormType, arrInTab), _
		"dialogWidth=380px; dialogHeight=200px; center: Yes; help: No; resizable: Yes; status: No; scrollbars: no")

	Set arrDoc(0) = Nothing

End Function
