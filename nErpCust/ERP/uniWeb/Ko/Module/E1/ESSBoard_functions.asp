<%
Function Tag2Text(text)
	Dim str
	str = Replace(text, "&" , "&amp;")
	str = Replace(str, "<", "&lt;")
	str = Replace(str, ">", "&gt;")
	Tag2Text = str
End Function

Function nLeft(str,strcut)
    Dim bytesize, nLeft_count
    bytesize = 0

    For nLeft_count = 1 to len(str)
        if asc(mid(str,nLeft_count,1)) > 0 then '한글값은 0보다 작다 
            bytesize = bytesize + 1 '한글이 아닌경우 1Byte
        else
            bytesize = bytesize + 2 '한글인 경우 2Byte
        end if
        if strcut >= bytesize then nLeft = nLeft & mid(str,nLeft_count,1)    
            '끊고싶은 길이(Byte)만큼 
    Next
 
	if  nLeft <> "" then
		if len(str) > len(nLeft) then nLeft= left(nLeft,len(nLeft)-2) & "..."
      '문자열이 짤렸을 경우 뒤에 ...을 붙여줌 
    end if
End Function

%>