<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<HTML>
<HEAD>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<STYLE>
.TH1
{
    BORDER-BOTTOM: buttonshadow 1px solid;
    BORDER-LEFT: buttonhighlight 1px solid; 
    BORDER-RIGHT: buttonshadow 1px solid;
    BORDER-TOP: buttonhighlight 1px solid
}
</STYLE>

<%
	'On Error Resume Next
	
	Dim strTable, strSQL 
	Dim oConn, oConn2, oRs, oRs2
	Dim intPageSize, intBlockPage, intTotalPage, intNowPage
	Dim prVarArray, intKeyNo	'�� �������ε��� ù��° Row�� ���� �۹�ȣ(NoticeNum)�� �о���� ����(frview.asp) �ʿ��� ���� 
	
	intPageSize = 10		'�� �������� �������� ����	
	intBlockPage = 10

	intNowPage = Request("page")

	If intNowPage = 0 Or Len(intNowPage) = 0 Then		
	    intNowPage = 1
	End If

	Call SubOpenDB(oConn)
	Call SubOpenDB(oConn2)
	
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	Set oRs2 = Server.CreateObject("ADODB.RecordSet")

	strSQL = "Select Count(*)"
	strSQL = strSQL & ",CEILING(CAST(Count(*) AS FLOAT)/" & intPageSize & ")"	
	strSQL = strSQL & " from B_NOTICE"

	Set oRs = oConn.Execute(strSQL,,1)

	intTotalCount = oRs(0)
	intTotalPage = oRs(1)
	
'	strSQL = "Select Top " & intNowPage * intPageSize & " A.NoticeNum"      & vbCr   ' ��ȣ 
'	strSQL = strSQL & ", MAX(A.Subject) Subject"                            & vbCr   ' ���� 
'	strSQL = strSQL & ", MAX(A.Writer) Writer"                              & vbCr   ' �̸� 
'	strSQL = strSQL & ", MAX(A.Usr_id) Usr_id"                              & vbCr   ' ���̵� 
'	strSQL = strSQL & ", MAX(A.RegDate) RegDate"                            & vbCr   ' ��¥ 
'	strSQL = strSQL & ", MAX(ISNULL(B.NoticeNum,'')) FILEYN "               & vbCr   ' ÷����������	
'	strSQL = strSQL & " FROM B_NOTICE A LEFT OUTER JOIN B_NOTICE_FILE B "   & vbCr
'	strSQL = strSQL & " ON A.NoticeNum = B.NoticeNum "		                & vbCr
'	strSQL = strSQL & " GROUP BY A.NoticeNum"	                            & vbCr
'	strSQL = strSQL & " ORDER BY A.NoticeNum desc"  

	strSQL = "Select Top " & intPageSize & " A.NoticeNum"      & vbCr   ' ��ȣ 
	strSQL = strSQL & ", A.Subject Subject"                            & vbCr   ' ���� 
	strSQL = strSQL & ", A.Writer Writer"                              & vbCr   ' �̸� 
	strSQL = strSQL & ", A.Usr_id Usr_id"                              & vbCr   ' ���̵� 
	strSQL = strSQL & ", A.RegDate RegDate"                            & vbCr   ' ��¥ 
	strSQL = strSQL & ", ISNULL(B.NoticeNum,'') FILEYN "               & vbCr   ' ÷����������	
	strSQL = strSQL & " FROM B_NOTICE A LEFT OUTER JOIN (SELECT NoticeNum from B_NOTICE_FILE GROUP BY NoticeNum) B "   & vbCr	
	strSQL = strSQL & " ON A.NoticeNum = B.NoticeNum "		                & vbCr
	strSQL = strSQL & " Where A.NoticeNum NOT IN ( SELECT TOP " & (intNowPage - 1) * intPageSize & " NoticeNum " & vbCr
	strSQL = strSQL & "                          FROM B_NOTICE "            & vbCr
	strSQL = strSQL & "                          ORDER BY NoticeNum Desc) "  & vbCr			
	strSQL = strSQL & " ORDER BY A.NoticeNum desc"  

	Set oRs = oConn.Execute(strSql,,1)

	If oRs.BOF Or oRs.EOF Then
		intKeyNo = 0	
	Else
'		prVarArray = oRs.GetRows()		
'		intKeyNo =  prVarArray(0,  (intNowPage - 1) * intPageSize)    
        intKeyNo = oRs(0)
'		oRs.MoveFirst		    
	End If

%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Script Language=VBScript FOR=myTitle EVENT=onclick>
	ChgNavyColor(window.event.srcElement)
</Script>

<Script Language="VBScript">
	Public MyTableRowIndex, intKeyNo, intNowPage, intPageSize, intTotalPage
	intNowPage = "<%=intNowPage%>"
	intKeyNo = "<%=intKeyNo%>"
	intPageSize = "<%=intPageSize%>"
	intTotalPage = "<%=intTotalPage%>"
	
	MyTableRowIndex = <%If oRs.EOF then Response.write -1 Else Response.write 0 End If %>
</Script>

<Script Language="VBScript">
	Function ChgNavyColor(wevent)
		Dim i, intNo, intUpNo, ev

		Select Case wevent.tagname
			Case "TD"
				Exit Function
			Case "FONT"
				Set ev = wevent.parentElement
			Case Else
				Exit Function
		End Select

		'msgbox "MyTableRowIndex=>"  & MyTableRowIndex & vbcrlf & vbcrlf & "ev.parentElement.rowindex=>" & ev.parentElement.rowindex

		If ev.parentElement.rowindex <> MyTableRowIndex Then		
				If MyTableRowIndex <> -1 Then 
					Call delReserve()	'���ڻ� �������� ����		
				End If
				
				MyTableRowIndex = ev.parentElement.rowindex		
				'intKeyNo = ev.parentElement.parentElement.rows(MyTableRowIndex).children(0).children(0).innerText				
				intKeyNo = ev.parentElement.parentElement.rows(MyTableRowIndex).children(1).children(0).innerText     'ù��° TD�� NoticeNum ���� ����÷������ Į������ �����. 
				parent.frView.location.href = "frView.asp?n=" & intKeyNo

		End If
		
		Call getfocus()		' ���� �ٲ� 

	End Function

	Sub window_onload()	

		myHead.style.borderColor = "black"
		myHead.style.borderBottom ="1 solid black"
		myHead.style.borderRight ="1 solid black"
		myHead.style.borderTop ="1 solid buttonhighlight"
		myHead.style.borderLeft = "1 solid buttonhighlight"

		Call getfocus
		
		if (myHead.offsetHeight + myTitle.offsetHeight + 5) > 100 Then
			parent.frMain.rows = (myHead.offsetHeight + myTitle.offsetHeight + 5) & ",*"
		End If

		parent.frView.location.href = "frView.asp?n=" & "<%=intKeyNo%>"
		
	End Sub

	Sub document_onclick()
		On Error Resume Next
		Call parent.frView.lostfocus()
	End Sub

	Sub getfocus()
		If MyTableRowIndex <> -1 Then
			myTitle.rows(MyTableRowIndex).bgcolor="FFF789"		'MSDN���� rows Collection ���� 
			For i = 0 to myTitle.rows(MyTableRowIndex).Cells.length-1
				myTitle.rows(MyTableRowIndex).children(i).children(0).color="black"
			Next
		End If
	End Sub

	Sub lostfocus()
		If MyTableRowIndex <> -1 Then
			myTitle.rows(MyTableRowIndex).bgcolor="F4F3F3"
			For i = 0 to myTitle.rows(MyTableRowIndex).Cells.length-1
				myTitle.rows(MyTableRowIndex).children(i).children(0).color="black"
			Next
		End If
	End Sub

	Sub delReserve()			'���ڻ� �������� ���� 
		myTitle.rows(MyTableRowIndex).bgcolor=""
		For i = 0 to myTitle.rows(MyTableRowIndex).Cells.length-1
			myTitle.rows(MyTableRowIndex).children(i).children(0).color="black"
		Next
	End Sub

	'========================================================================================
	' Function Name : Document_onKeyDown
	' Function Desc : hand all event of key down
	'========================================================================================
	Function Document_onKeyDown()
		Dim objEl, KeyCode, iLoc
		Dim boolMinus, boolDot
		Document_onKeyDown = True
		Set objEl = window.event.srcElement
		KeyCode = window.event.keycode
	    Set gActiveElement = document.activeElement
		Select Case KeyCode	
			Case 123  'F12
				Window.top.Frames(1).Focus
				Window.top.Frames(1).SetMenuHightLight(Window.top.Frames(1).gCurP)
				Window.top.Frames(1).gF12KeyEnable = True
				Document_onKeyDown = False
				Exit Function	
		End Select	
	End Function

	'========================================================================================
	' ���� form_load �Լ��� ���� ������ ������!!!!!
	'========================================================================================
	Sub form_load()
	    gFocusSkip = True
	End Sub
</Script>

<!-- #Include file="../inc/UNI2KCMCom.inc" -->	

</HEAD>
<BODY>

<% If intTotalCount > 0 Then %>
<!--table width="600">
  <TR>
    <TD>��ü�Խ� <%=intTotalCount%> �� &nbsp;&nbsp;&nbsp;&nbsp;
            ���������� : <%=intNowPage%> / <%=intTotalPage%>
    </TD>
  </TR>
</table-->
<%  End If  %>
<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0 bgcolor="#F4F3F3">
	<TR>
		<TD WIDTH=100% height=20>
			<TABLE ID="myHead" NAME="myHead" WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=0>
				<TR BGCOLOR="D1E8F9">
					<TH WIDTH="3%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>&nbsp;</FONT></TD>									
					<TH STYLE="display:none;" WIDTH="7%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>No</FONT></TD>
					<TH WIDTH="20%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>�ۼ���</FONT></TD>
					<TH WIDTH="10%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>���̵�</FONT></TD>
					<TH WIDTH="50%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>����</FONT></TD>
					<TH WIDTH="17%" BORDERCOLOR=black CLASS="TH1" ALIGN=center>&nbsp;<FONT FACE="<%=g33FontName%>" size=2>�ۼ���</FONT></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD height=* valign=top>	
		<TABLE ID=myTitle Name="myTitle" WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1 bgcolor="#F4F3F3">		
		<%  If oRs.BOF or oRs.EOF Then 		
				Response.Write  "��ϵ� �Խù��� �������� �ʽ��ϴ�."
		    Else
		            'oRs.Move (intNowPage - 1) * intPageSize
		            Do Until oRs.EOF		
		%>
			<TR>
				<%
				'lgStrSQL = "SELECT 1 FROM B_NOTICE_FILE WHERE NOTICENUM = " & oRs("NoticeNum") 
				'Set oRs2 = oConn2.Execute(lgStrSQL,,1)
				'If oRs2.BOF Or oRs2.EOF Then
			    If Trim(oRs("FILEYN")) = "0" Then '÷�������� ���� ��� 
				%>
				<TD WIDTH=3% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2>&nbsp;</TD>							
				<%				
				Else
				%>
				<TD WIDTH=3% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2><img src=../../CShared/image/clip1.gif width=10 height=13 border=0></font></TD>
				<%				
				End If
				%>
				<TD STYLE="display:none;" WIDTH=7% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2><%=oRs("NoticeNum")%><input type=hidden name="txtNo" value="<%=oRs("NoticeNum")%>"></font></TD>
				<TD WIDTH=20% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2><%=oRS("Writer")%></font></TD>
				<TD WIDTH=10% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2><% If ""&oRs("Usr_id") = "" Then%>��ü<% Else %><%= oRS("Usr_id") %><% End If %></font></TD>
				<TD WIDTH=50% ALIGN=LEFT><font FACE="<%=g33FontName%>" size=2><%= oRS("Subject") %></font></TD>
				<TD WIDTH=17% ALIGN=CENTER><font FACE="<%=g33FontName%>" size=2><%=oRS("RegDate")%></font></TD>
			</TR>
		<%
						oRs.MoveNext 
					Loop
		%>

			<TR align="center">
			  <TD colspan=5>
			  <%
					  
			          intTemp = Int((intNowPage - 1) / intBlockPage) * intBlockPage + 1			          
						'������ intTemp ��� ������, ������ ��� ����� [���� 10��] �� [���� 10��]�� ��ũ�� Ŭ������ ��� 
						'�������� ù��° ������ (1������, 11������, 21������, 31������...)�� ����ϱ� ���� �ӽ����� �����Դϴ�.
						'���� intBlockPage�� ���� 5�� �������� ��� intTemp�� ���� 1, 6, 11, 16...�� ��.
						
			          If intTemp = 1 Then
			              'Response.Write "[���� " & intBlockPage & "��]&nbsp;"
			              Response.Write "<img src=../../CShared/image/arrow/left2_deactivated.gif width=14 height=13 border=0 alt=""" & "���� " &intBlockPage & " ���� ��&nbsp;" & """>&nbsp;"
			          Else 
			              'Response.Write"<a href=frtitle.asp?page=" & intTemp - intBlockPage & ">[���� " & intBlockPage & "��]</a>&nbsp;"
						  Response.Write"<A href=frtitle.asp?page=" & intTemp - intBlockPage & "><img src=../../CShared/image/arrow/left2_activated.gif width=14 height=13 border=0 alt=""" & "���� " &intBlockPage & " ���� ��&nbsp;" & """></A>&nbsp;" 					  			              
			          End If

					  If intNowPage = 1 Then
						  Response.Write"<img src=../../CShared/image/arrow/left_deactivated.gif width=14 height=13 border=0 alt=����>&nbsp;"					  
					  Else
						  Response.Write"<A href=frtitle.asp?page=" & intNowPage - 1 & "><img src=../../CShared/image/arrow/left_activated.gif width=14 height=13 border=0 alt=����></A>&nbsp;" 					  
					  End If
					  
			          intLoop = 1

			          Do Until intLoop > intBlockPage Or intTemp > intTotalPage
			              If intTemp = CInt(intNowPage) Then
			                  Response.Write "[<font size= 3><b>" & intTemp &"</b></font>]&nbsp;" 
			              Else
			                  Response.Write"[<a style=""color=#2328FA"" href=frtitle.asp?page=" & intTemp & ">" & intTemp & "</a>]&nbsp;"
			              End If
			              intTemp = intTemp + 1
			              intLoop = intLoop + 1
			          Loop
			          
    				'Response.Write "<Script Language=vbscript>"            & vbCr
    				'Response.Write "msgbox """   & intNowPage &  """"     & vbCr
					'Response.Write "msgbox """   & intTemp &  """"     & vbCr
					'Response.Write "msgbox """   & intTotalPage &  """"     & vbCr	
					'Response.Write "</Script>"                             & vbCr
			          
					  If CInt(intNowPage) = Cint(intTotalPage) Then
                          Response.Write"<img src=../../CShared/image/arrow/right_deactivated.gif width=14 height=13 border=0 alt=����>&nbsp;"
					  Else
						  Response.Write"<A href=frtitle.asp?page=" & intNowPage + 1 & "><img src=../../CShared/image/arrow/right_activated.gif width=14 height=13 border=0 alt=����></A>&nbsp;"
					  End If			          

			              
			          If intTemp > intTotalPage Then
			              'Response.Write "[���� " &intBlockPage&"��]&nbsp;"			              
			              Response.Write"<img src=../../CShared/image/arrow/right2_deactivated.gif width=14 height=13 border=0 alt=""" & "���� " &intBlockPage & " ���� ��&nbsp;" & """>&nbsp;"
			          Else
			              'Response.Write"<a href=frtitle.asp?page=" & intTemp & ">[���� " & intBlockPage & "��]</a>&nbsp;"
			              Response.Write"<A href=frtitle.asp?page=" & intTemp & "><img src=../../CShared/image/arrow/right2_activated.gif width=14 height=13 border=0 alt=""" & "���� " &intBlockPage & " ���� ��&nbsp;" & """></A>&nbsp;"
			          End If
			  %>
			  </TD>
			</TR>
		
		<%
			End If

			oRs.Close
			oConn.Close 
			set oRs = Nothing
			set oConn = Nothing
		%>		
		</TABLE>
</TABLE>
</BODY>
</HTML>
