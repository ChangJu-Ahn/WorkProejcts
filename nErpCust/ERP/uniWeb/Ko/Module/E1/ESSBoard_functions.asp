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
        if asc(mid(str,nLeft_count,1)) > 0 then '�ѱ۰��� 0���� �۴� 
            bytesize = bytesize + 1 '�ѱ��� �ƴѰ�� 1Byte
        else
            bytesize = bytesize + 2 '�ѱ��� ��� 2Byte
        end if
        if strcut >= bytesize then nLeft = nLeft & mid(str,nLeft_count,1)    
            '������� ����(Byte)��ŭ 
    Next
 
	if  nLeft <> "" then
		if len(str) > len(nLeft) then nLeft= left(nLeft,len(nLeft)-2) & "..."
      '���ڿ��� ©���� ��� �ڿ� ...�� �ٿ��� 
    end if
End Function

%>