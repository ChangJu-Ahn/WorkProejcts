 <%
 
Sub GoToPageDirectly(page, Pagecount)
'Response.Write "page:" & page
'Response.Write ",Pagecount:" & Pagecount
	
	'myBlodkEnd�� ���� �ڽ��� �������� ��ġ�ϴ� ���� ������ ���̴�.
	'��, ���� �������� 3�̶�� myBlodkEnd�� 10�� �ǰ�..
	'���� �������� 10 �̾ myBlodkEnd �� 10�� �ȴ�.
	'����, ���� �������� 11�̶�� myBlodkEnd �� 20�� �ǰ�...
	'���� �������� 20 �̶�� 20�� �ȴ�.
	
    Dim myBlodkEnd 
    Dim endNum : endNum = Right(page, 1)
    
    '���� �ڽ��� ������ ������ ������ ������ ���ϱ�.
    If (page Mod 10) = 0 Then
        myBlodkEnd = page
    Else
        myBlodkEnd = (Int(page) + 10) - Int(endNum)   '13 + 10 - 3  / 23 + 10 -3
    End If
'Response.Write ",myBlodkEnd:" & myBlodkEnd
    '���� 10�� ��� ���� 
    If Int(myBlodkEnd) > 10 Then
        Response.Write "<a href='List.asp?page=" & myBlodkEnd-19 & "'><img src=""../../CShared/EISImage/ENotice/ic_mp_first.gif"" align=absmiddle border=0></a>"
    Else
		Response.Write " <img src=""../../CShared/EISImage/ENotice/ic_mp_first.gif"" align=absmiddle border=0> "
    End If
    

    '���� �������� ���� ��� ���� 
    if int(right(page,1)) = 1 then
		Response.Write " <img src=../../CShared/EISImage/ENotice/ic_mp_prev.gif border=0 align=absmiddle>&nbsp; " 
    else
        Response.Write " <a href='List.asp?page=" & page - 1 & "'><img src=../../CShared/EISImage/ENotice/ic_mp_prev.gif border=0 align=absmiddle></a> &nbsp;" 
    end if
    
    
    Dim i, endNumOfLoop
    If Int(pagecount) > Int(myBlodkEnd) Then
		endNumOfLoop = myBlodkEnd
	else
		endNumOfLoop = Int(pagecount)
	end if
	
    For i = myBlodkEnd - 9 To endNumOfLoop
		if i = int(page) then 
			Response.Write "<font style='color:silver'>" & i & "</font>"
		else
			Response.Write " <a href='List.asp?page=" & i & "'>" & i & "</a> " 
		end if
    Next
    
    '���� �������� ���� ��� ���� 
    if int(page) = endNumOfLoop then
		Response.Write " &nbsp;<img src=../../CShared/EISImage/ENotice/ic_mp_next.gif border=0 align=absmiddle> " 
    else
        Response.Write " &nbsp;<a href='List.asp?page=" & page + 1 & "'><img src=../../CShared/EISImage/ENotice/ic_mp_next.gif border=0 align=absmiddle></a> " 
    end if
    
    '���� 10�� ��� ���� 
    If Int(pagecount) > Int(myBlodkEnd) Then
        Response.write " <a href='List.asp?page=" & myBlodkEnd+1 & "'><img src=""../../CShared/EISImage/ENotice/ic_mp_last.gif"" align=absmiddle border=0></a>"
    else
		Response.Write " <img src=""../../CShared/EISImage/ENotice/ic_mp_last.gif"" align=absmiddle border=0> " 
    End If
    
End Sub


Sub GotoPageInSearchResult(page, Pagecount, SearchPart, SearchStr)
    
    Dim myBlodkEnd 
    Dim endNum : endNum = Right(page, 1)
    
    '���� �ڽ��� ������ ������ ������ ������ ���ϱ�.
    If (page Mod 10) = 0 Then
        myBlodkEnd = page
    Else
        myBlodkEnd = (Int(page) + 10) - Int(endNum)   '13 + 10 - 3  / 23 + 10 -3
    End If

    '���� 10�� ��� ���� 
    If Int(myBlodkEnd) > 10 Then
        Response.Write "<a href='Search.asp?page=" & myBlodkEnd-19 & "&table=" & table & "&SearchPart=" & SearchPart & "&SearchStr=" & SearchStr & "'><img src='../../CShared/EISImage/ENotice/ic_mp_first.gif' align=absmiddle border=0></a>"
    Else
		Response.Write " <img src=""../../CShared/EISImage/ENotice/ic_mp_first.gif"" align=absmiddle border=0> "
    End If
    

    '���� �������� ���� ��� ���� 
    if int(right(page,1)) = 1 then
		Response.Write " <img src=../../CShared/EISImage/ENotice/ic_mp_prev.gif border=0 align=absmiddle>&nbsp; " 
    else
        Response.Write " <a href='Search.asp?page=" & page - 1 & "&table=" & table & "&SearchPart=" & SearchPart & "&SearchStr=" & SearchStr & "'><img src=../../CShared/EISImage/ENotice/ic_mp_prev.gif border=0 align=absmiddle></a> &nbsp;" 
    end if
    
    
    Dim i, endNumOfLoop
    If Int(pagecount) > Int(myBlodkEnd) Then
		endNumOfLoop = myBlodkEnd
	else
		endNumOfLoop = Int(pagecount)
	end if
	
    For i = myBlodkEnd - 9 To endNumOfLoop
		if i = int(page) then 
			Response.Write "<font style='color:silver'>" & i & "</font>"
		else
			Response.Write " <a href='Search.asp?page=" & i & "&table=" & table & "&SearchPart=" & SearchPart & "&SearchStr=" & SearchStr & "'>" & i & "</a> " 
		end if
    Next
    
    '���� �������� ���� ��� ���� 
    if int(page) = endNumOfLoop or endNumOfLoop = 0 then
		Response.Write " &nbsp;<img src=../../CShared/EISImage/ENotice/ic_mp_next.gif border=0 align=absmiddle> " 
    else
        Response.Write " &nbsp;<a href='Search.asp?page=" & page + 1 & "&table=" & table & "&SearchPart=" & SearchPart & "&SearchStr=" & SearchStr & "'><img src=../../CShared/EISImage/ENotice/ic_mp_next.gif border=0 align=absmiddle></a>" 
    end if
    
    '���� 10�� ��� ���� 
    If Int(pagecount) > Int(myBlodkEnd) Then
        Response.write " <a href='Search.asp?page=" & myBlodkEnd+1 & "&table=" & table & "&SearchPart=" & SearchPart & "&SearchStr=" & SearchStr & "'><img src=../../CShared/EISImage/ENotice/ic_mp_last.gif align=absmiddle border=0></a>"
    else
		Response.Write " <img src=../../CShared/EISImage/ENotice/ic_mp_last.gif align=absmiddle border=0> " 
    End If

End Sub
%>
