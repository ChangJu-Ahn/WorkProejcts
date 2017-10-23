Sub CalculateMath(ByRef txtClientNum1000, ByRef txtClientNumDec, ByRef txtClientDateFormat, ByRef txtClientDateSeperator)

    Dim pClientDateFormat
    Dim pClientDateSeperator
 
    Dim iLoop

    Dim pClientNumDec
    Dim pClientNum1000
    
    Dim Tmp

    pClientDateFormat = DateAdd("D", -1, "2004-01-01")

    pClientDateFormat = Replace(pClientDateFormat, "2003", "YYYY")
    pClientDateFormat = Replace(pClientDateFormat, "03", "YY")
    pClientDateFormat = Replace(pClientDateFormat, "Dec", "MMM")
    pClientDateFormat = Replace(pClientDateFormat, "12", "MM")
    pClientDateFormat = Replace(pClientDateFormat, "31", "DD")

    For iLoop = 1 To Len(pClientDateFormat)
        pClientDateSeperator = Mid(pClientDateFormat, iLoop, 1)
        If Not (pClientDateSeperator = "Y" Or pClientDateSeperator = "M" Or pClientDateSeperator = "D") Then
           Exit For
        End If
    Next

    Tmp = FormatNumber(12345 / 10, 2, -1, -1, -1)
    
    Tmp = FormatNumber(12345 / 10, 2, -1, -1, -1)

    pClientNum1000 = Mid(Tmp, 2, 1) ' 1,234.50
    pClientNumDec = Mid(Tmp, 6, 1)

    txtClientNum1000 = pClientNum1000
    txtClientNumDec = pClientNumDec
    txtClientDateFormat = pClientDateFormat
    txtClientDateSeperator = pClientDateSeperator

End Sub

Sub WebPathInfo(ByRef pHTTP,ByRef pServer,ByRef pVD,ByRef pLang)

    Dim iTemp
    Dim iTempArr

	iTemp = window.location.href
	
	iTempArr = Split(iTemp,"/")

	pHTTP   = iTempArr(0)
	pServer = iTempArr(2)
	pVD     = iTempArr(3)
	pLang   = UCase(iTempArr(4))

End Sub