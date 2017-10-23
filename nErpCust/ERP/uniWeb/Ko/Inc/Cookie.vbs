Function WriteCookie(varCookie, varValue)	
        varValue = Replace(varValue,";","<:::>") 
	Document.Cookie = varCookie & "=" & Escape(varValue) & "; path=" & "/"
End Function

Function WriteExpiresCookie(varCookie, varValue, ExpiresValue)	

	Dim MyWeek, MyMonth, MyDay, ExpDate
	
	MyDay = Date()
	MyDay = MyDay + 5
	MyMonth = Month(MyDay)
	If Len(MyMonth) = 1 Then MyMonth = "0" & MyMonth
	MyWeek = Array( "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
	
	ExpDate = MyWeek(WeekDay(MyDay) - 1) & ", " & Day(MyDay) & "-" & MyMonth & "-" & Year(MyDay) & " 12:00:00 GMT"
	Document.Cookie = varCookie & "=" & Escape(varValue) & "; path=" & "/; expires=" & ExpDate
	
End Function
	
Function ReadCookie(varCookie)
	Dim iArrCookie
	Dim iCookievalue
	Dim i
	Dim iName, iValue, iPos	

	iArrCookie = Split(Document.Cookie,";")
	
	ReadCookie = ""
	For i = 0 To UBound(iArrCookie)
	    iPos = Instr(iArrCookie(i),"=")
	    If iPos <> 0 Then
	        iName = Trim(left(iArrCookie(i) , iPos - 1))
	        iValue = Trim(Mid(iArrCookie(i) , iPos +1))
	    Else
	        iName = Trim(iArrCookie(i))
	        iValue = ""
	    End If
	    
	    If iName = Trim(varCookie) Then
	        ReadCookie = UnEscape(iValue)
            ReadCookie = Replace(ReadCookie  ,"<:::>",";") 
            Exit Function
	    End If
	Next
End Function
