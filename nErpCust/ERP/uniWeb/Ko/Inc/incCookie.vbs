'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' <<<<<<<< Cookie 관련 함수>>>>>>>>
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'========================================================================================================
' Name : WriteCookie(varCookie, varValue)	
' Desc : WriteCookie
'========================================================================================================

Function WriteCookie(varCookie, varValue)	
	Document.Cookie = varCookie & "=" & varValue & ";"
End Function

'========================================================================================================
' Name : ExpiresCookie(varCookie)	
' Desc : ExpiresCookie
'========================================================================================================
Function ExpiresCookie(varCookie)

	Dim MyWeek, MyMonth, MyDay, ExpDate
	
	varCookie = Replace(varCookie,"_","%5F")
	varCookie = Replace(varCookie," ","+")
	varCookie = "unierp" & Chr(38) & varCookie
	MyDay = Date()
	MyDay = MyDay - 1
	MyMonth = Month(MyDay)
	If Len(MyMonth) = 1 Then MyMonth = "0" & MyMonth
	MyWeek = Array( "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
	
	ExpDate = MyWeek(WeekDay(MyDay) - 1) & ", " & Day(MyDay) & "-" & MyMonth & "-" & Year(MyDay) & " 12:00:00 GMT"
	Document.Cookie = varCookie & "=; expires=" & ExpDate

End Function
	
'========================================================================================================
' Name : ReadCookie(varCookie)
' Desc : ReadCookie
'========================================================================================================
Function ReadCookie(varCookie)
	Dim intLocation
	Dim intNameLength
	Dim intValueLength
	Dim intNextSemicolon
	Dim strTemp
	ReadCookie = ""

	varCookie = Replace(varCookie,"_","%5F")
	varCookie = Replace(varCookie," ","+")		
	intNameLength = Len(varCookie)
	intLocation = Instr(Document.Cookie, varCookie)
 
	If intLocation = 0 Then
		ReadCookie = ""
	Else		
		strTemp = Right(Document.Cookie, Len(Document.Cookie) - intLocation + 1)
		If Mid(strTemp, intNameLength + 1, 1) <> "=" Then	 
	 		ReadCookie = ""
		Else
	 		intNextSemicolon = Instr(strTemp, "&")
			If intNextSemicolon = 0 Then intNextSemicolon = Instr(strTemp, ";")
			If intNextSemicolon = 0 Then intNextSemicolon = Len(strTemp) + 1
			
			If intNextSemicolon = (intNameLength + 2) Then	
	     		ReadCookie = ""
	 		Else	     
     			intValueLength = intNextSemicolon - intNameLength - 2
     			ReadCookie = ConvCookieSPChars(Replace(Mid(strTemp, intNameLength + 2, intValueLength),"+"," "))
	 		End If
		End If
	End If
End Function
'========================================================================================
' Function Name : ConvCookieSPChars
' Function Desc : 쿠키값을 decoding
'========================================================================================
Function ConvCookieSPChars(iStr)
    Dim ii 
    Dim iEnCD
    ConvCookieSPChars = ""     
    ii = 1
    Do While ii <= Len(iStr)
        If Mid(iStr,ii,1) = "%" Then
            iEnCD = ""
            Select Case Mid(iStr,ii,3)
               Case "%21"
                         iEnCD = "!"
               Case "%23"
                         iEnCD = "#"
               Case "%24"
                         iEnCD = "$"
               Case "%25"
                         iEnCD = "%"
               Case "%26"
                         iEnCD = "&"
               Case "%27"
                         iEnCD = "'"
               Case "%28"
                         iEnCD = "("
               Case "%29"
                         iEnCD = ")"
               Case "%2B"
                         iEnCD = "+"
               Case "%2D"
                         iEnCD = "-"
               Case "%2E"
                         iEnCD = "."
               Case "%2F"
                         iEnCD = "/"
               Case "%3A"
                         iEnCD = ":"
               Case "%3B"
                         iEnCD = ";"
               Case "%3C"
                         iEnCD = "<"
               Case "%3D"
                         iEnCD = "="
               Case "%3E"
                         iEnCD = ">"
               Case "%3F"
                         iEnCD = "?"
               Case "%5B"
                         iEnCD = "["
               Case "%5C"
                         iEnCD = "\"
               Case "%5D"
                         iEnCD = "]"
               Case "%5E"
                         iEnCD = "^"
               Case "%5F"
                         iEnCD = "_"
               Case "%60"
                         iEnCD = "`"
               Case "%7B"
                         iEnCD = "{"
               Case "%7C"
                         iEnCD = "|"
               Case "%7D"
                         iEnCD = "}"
               Case "%0D"
                         iEnCD = Chr(13)
            End Select              
            ConvCookieSPChars = ConvCookieSPChars &  iEnCD
            ii = ii + 3
       Else
            ConvCookieSPChars = ConvCookieSPChars &  Mid(iStr,ii,1)
            ii = ii + 1
       End If
Loop
End Function
