<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->


<HTML>
<HEAD>
<STYLE type='text/css'> 
	A
	{
	    COLOR: black;
	    FONT-SIZE: 10pt;
	    TEXT-DECORATION: none
	}
	
	A:link    {color:black;font-size:9pt;text-decoration:none;}
	A:visited {color:black;font-size:9pt;text-decoration:none;}
	A:active  {color:black;font-size:9pt;text-decoration:none;}
	A:hover  {color:#008400;text-decoration:underline;}
	
	.TH1
	{
	    BORDER-BOTTOM: 1 solid black;
	    BORDER-RIGHT: 1 solid black;
	    BORDER-TOP: 1 solid black;
	}	

	.TH2
	{
	    BORDER-LEFT: 1 solid black;	
	}		

	.TH3
	{
	    BORDER-BOTTOM: 1 solid black;
	    BORDER-LEFT: 1 solid black;
	    BORDER-RIGHT: 1 solid black;
	    BORDER-TOP: 1 solid black;
	}	
	
</STYLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%

'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    

Dim lgObjConn, lgObjComm,lgObjRs, lgObjRs2, lgObjRs3, lgStrSQL
Dim strKeyNO, blnStop
Dim strWriter, strTitle, strContent, strUsr_id, strRegDate

Dim strData
Dim arrFileInfo
Dim iDx
Dim iCount
Dim tmpString
Dim tmpString1
Dim tmpString2
Dim iTempPath

Dim fso
Dim strSystemFolder

If Not (IsEmpty(Request("n")) Or Request("n") = "") Then
	strKeyNO = CLng(Request("n"))
Else
	strKeyNO = 0
End If

strWriter = "" : strTitle = "": strContent = "": strUsr_id = "" : strRegDate = ""
blnStop = False

If strKeyNO <> 0 Then

	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)	
	
	' SQL문 작성 
	lgStrSQL = "SELECT Writer, Subject, Usr_id, Contents, RegDate, (SELECT COUNT(FLE_ID) FROM B_NOTICE_FILE WHERE NOTICENUM = " & strKeyNO & ") AttachedFileNum FROM B_NOTICE WHERE NoticeNum=" & strKeyNO
	'lgStrSQL = "SELECT Writer, Subject, Usr_id, Contents0, RegDate FROM B_NOTICE WHERE ROWNUM = 1 AND NoticeNum=" & strKeyNO 'FOR HERMES

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
		Call DisplayMsgBox("210037", vbInformation, "", "", I_MKSCRIPT)      '☜ : 게시물에 해당하는 자료가 존재하지 않습니다.

		Response.Write "<Script Language=vbscript>"			& vbCr
		Response.write "	Parent.window.frames(""frTitle"").document.URL = ""frtitle.asp?page=""" & "& CStr(parent.window.frames(""frTitle"").intNowPage)" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End					
	Else
		iCount = lgObjRs("AttachedFileNum")
	End If

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If
    	
	Redim arrFileInfo(100)
		
	lgStrSQL = "SELECT FLE_NAME, FLE_ID, FLE_PATH FROM B_NOTICE_FILE WHERE NOTICENUM = " & strKeyNO & " ORDER BY INSRT_DT ASC"


	strSystemFolder = GetSpecialFolder(0) '0->WindowsFolder, 1->SystemFolder, 2->TemporaryFolder		
	strSystemFolder = strSystemFolder & "\TEMP"
	
	
	If right(strSystemFolder,1) <> "\" Then
		iTempPath = strSystemFolder & "\UNIERPTEMP\"
	Else
		iTempPath = strSystemFolder & "UNIERPTEMP\"
	End If	

	'TEMP폴더 없으면 생성 
	Call CreateFolder(iTempPath)
	
	iTempPath = Replace(iTempPath, "\", "/")
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs2,lgStrSQL,"X","X") = False Then                    'If data not exists	
	Else
		iDx = 0
		Do While Not lgObjRs2.EOF		
			'strData = "1" & lgObjRs2(0) & "F:/work/patch/Board/mySingle/ocx/TEMP/" & lgObjRs2(0) & "" & lgObjRs2(2) & lgObjRs2(1) & "101344919970601092656N0NNFFY00YI"
			strData = "1" & lgObjRs2(0) & "" & lgObjRs2(2) & lgObjRs2(0) & "" & iTempPath & lgObjRs2(1) & "101344919970601092656N0NNFFY00YI"										
			'DisplayMessageBox(strData)
			arrFileInfo(iDx) = strData
			lgObjRs2.MoveNext
			iDx = iDx + 1
		Loop
	End If

	If Not lgObjRs.EOF Then	
		
		strWriter = "" & lgObjRs("Writer")
		strTitle = "" & lgObjRs("Subject")
		
		strRegDate = "" & lgObjRs("RegDate")
		strContent = "" & lgObjRs("Contents")
		'strContent = "" & lgObjRs("Contents0")   'FOR HERMES		
		strUsr_id = "" & lgObjRs("Usr_id")

%>

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>

<Script Language="VBScript">

Const BIZ_PGM_ID = "frviewBiz.asp"												'☆: Head Query 비지니스 로직 ASP명 

Sub window_onload()
	myView.style.borderBottom ="1 solid black"
	myView.style.borderRight ="1 solid black"
	myView.style.borderTop ="1 solid buttonhighlight"
	myView.style.borderLeft = "1 solid buttonhighlight"
End Sub

Sub SetFileViewStyle()
	FileView.style.borderBottom ="1 solid black"
	FileView.style.borderRight ="1 solid black"
	FileView.style.borderTop ="1 solid buttonhighlight"
	FileView.style.borderLeft = "1 solid buttonhighlight"			
End Sub

Sub window_onblur()
	Call lostfocus
End Sub

Sub document_onclick()
	Call getfocus
End Sub

Function GetFile(iDx)

	Select Case iDx
		Case 0
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(0)%>"
		Case 1
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(1)%>"
		Case 2
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(2)%>"
		Case 3
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(3)%>"
		Case 4
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(4)%>"
		Case 5
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(5)%>"
		Case 6
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(6)%>"
		Case 7
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(7)%>"
		Case 8
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(8)%>"
		Case 9
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(9)%>"
		Case 10
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(10)%>"
		Case 11
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(11)%>"
		Case 12
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(12)%>"
		Case 13
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(13)%>"
		Case 14
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(14)%>"
		Case 15
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(15)%>"
		Case 16
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(16)%>"
		Case 17
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(17)%>"
		Case 18
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(18)%>"
		Case 19
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(19)%>"
		Case 20
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(20)%>"
		Case 21
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(21)%>"
		Case 22
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(22)%>"
		Case 23
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(23)%>"
		Case 24
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(24)%>"
		Case 25
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(25)%>"
		Case 26
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(26)%>"
		Case 27
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(27)%>"
		Case 28
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(28)%>"
		Case 29
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(29)%>"
		Case 30
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(30)%>"
		Case 31
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(31)%>"
		Case 32
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(32)%>"
		Case 33
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(33)%>"
		Case 34
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(34)%>"
		Case 35
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(35)%>"
		Case 36
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(36)%>"
		Case 37
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(37)%>"
		Case 38
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(38)%>"
		Case 39
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(39)%>"
		Case 40
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(40)%>"
		Case 41
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(41)%>"
		Case 42
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(42)%>"
		Case 43
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(43)%>"
		Case 44
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(44)%>"
		Case 45
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(45)%>"
		Case 46
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(46)%>"
		Case 47
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(47)%>"
		Case 48
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(48)%>"
		Case 49
			strVal = BIZ_PGM_ID & "?FileInfo=" & "<%=arrFileInfo(49)%>"
																								
	End Select			
	
	strVal = strVal & "&FileMode=" & Trim(cboFileMode.value)
	
	Call RunMyBizASP(MyBizASP, strVal)
		
End Function

Sub getfocus()
	myView.bgColor = "navy"
	For i = 0 to myFnt.length-1
		myFnt(i).color = "white"
	Next
	Call Parent.frTitle.lostfocus
End Sub

Sub lostfocus()
	myView.bgColor = "D1E8F9"
	For i = 0 to myFnt.length-1
		myFnt(i).color = "black"
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
' 다음 form_load 함수를 절대 지우지 마세요!!!!!
'========================================================================================
Sub form_load()
    gFocusSkip = True
End Sub

Function vbCheckFileAssociation(sExt)
	vbCheckFileAssociation = CheckFileAssociation(sExt)	
End Function

Function vbViewFile(sMode, sRet)
	Call ViewFile(sMode, sRet)
End Function

Function FetchWebSvrIp()
	
	Dim gHttpWebSvrIPURL
	
	gHttpWebSvrIPURL = "http://<%= request.servervariables("server_name") %>"	
	FetchWebSvrIp    = Split(gHttpWebSvrIPURL, "/")(2)

End Function
</Script>

<script language=javascript>
function CheckFileAssociation(sExt){
	return document.FR_ATTWIZ.CheckFileAssociation(sExt);	
}
	
function ViewFile(sMode, sRet){
	var temp;
	var strWebSvrIp;
	
	document.FR_ATTWIZ.SetLanguage('K');	
	document.FR_ATTWIZ.SetModUpload();
	document.FR_ATTWIZ.SetServerAutoDelete(1);
	document.FR_ATTWIZ.SetFileUIMode(1);
	//document.FR_ATTWIZ.SetExtension('gultxt');
	//document.FR_ATTWIZ.SetServerOption(0,1);
	document.FR_ATTWIZ.SetServerOption(0,0);

	//웹서버ip Fetch
	document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	document.FR_ATTWIZ.ViewFile(sMode, sRet);
}	
</script>

<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>

<BODY bgcolor="#F4F3F3" topmargin=0>
<TABLE ID="myView" Name="myView" WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="D1E8F9">
	<TR>
		<TD>&nbsp;<font FACE="<%=g33FontName%>" size=2 ID="myFnt"><b>작성자: </b><%= strWriter  %></font></TD>
		<TD>&nbsp;<font FACE="<%=g33FontName%>" size=2 ID="myFnt"><b>제목: </b><%= strTitle  %></font></TD>
	</TR>
		<TR>
		<TD>&nbsp;<font FACE="<%=g33FontName%>" size=2 ID="myFnt"><b>아이디: </b><%= strUsr_id  %></font></TD>
		<TD>&nbsp;<font FACE="<%=g33FontName%>" size=2 ID="myFnt"><b>작성일: </b><%= strRegDate  %></a></font></TD>
	</TR>
</TABLE>

<%

	lgStrSQL = "SELECT FLE_NAME FROM B_NOTICE_FILE WHERE NOTICENUM = " & strKeyNO & " ORDER BY INSRT_DT ASC"
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs3,lgStrSQL,"X","X") = False Then                    'If data not exists	
%>
<TABLE height="90%" width="100%" bgcolor="#F4F3F3">
	<TR>
		<TD valign="top" style="zoom:100%;word-break:break-all"><%= Replace(strContent, chr(13), "<BR>") %></TD>
	</TR>
</TABLE>
<%	
	Else			
%>

<TABLE height="65%" width="100%" bgcolor="#F4F3F3">
	<TR>
		<TD valign="top" style="zoom:100%;word-break:break-all"><%= Replace(strContent, chr(13), "<BR>") %></TD>
	</TR>
</TABLE>

<TABLE ID="FileView" Name="FileView" HEIGHT=25% WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="D1E8F9" CLASS=TH1>
		<!--TR><TD width="100%" colspan=2>&nbsp;<font FACE="<%=g33FontName%>" size=2><b>첨부파일:</b></font></TD></TR-->
		<TR>
			<TD width="100%" colspan=2>
				<table width="100%" WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="D1E8F9" CLASS=TH1>
				  <tr>
				    <td>&nbsp;<font FACE="<%=g33FontName%>" size=2><b>첨부파일:</b></font></td>
				  </tr>
				</table>				
			</TD>
		</TR>
		<TR>
			<TD width="20%">			
				&nbsp;<font FACE="<%=g33FontName%>" size=2><b>실행모드:</b></font>
				&nbsp;&nbsp;<SELECT NAME=cboFileMode><OPTION value="W">저장</OPTION><OPTION value="F" selected>보기</OPTION></SELECT>			
			</TD>
			<TD width="80%">
				<div id="divDetachedFileInf"></div>						
				<%
	'			Response.Write "<Script Language=vbscript>"            & vbCr
	'			Response.Write "    Call SetFileViewStyle() "
	'			Response.Write "</Script>"                             & vbCr

				If lgObjRs3.BOF or lgObjRs3.EOF Then 		
					'Response.Write  "첨부된 파일이 없습니다."
				Else
						
					'첨부된 파일 정보를 테이블로 조립하는 부분 
					Dim iOutStr
					iDx = 0
					iOutStr = "<TABLE ID=FileList Name=FileList WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor=D1E8F9 CLASS=TH2>"
					iOutStr = iOutStr & "<TR><TD>&nbsp;</TD></TR>"
				    Do Until lgObjRs3.EOF		
						iOutStr = iOutStr & "<TR><TD>&nbsp;<A onclick='VBSCRIPT:GETFILE(" & iDx & ")'>" & lgObjRs3(0) & "</A></TD></TR>"
						lgObjRs3.MoveNext 
						iDx = iDx + 1
					Loop
					iOutStr = iOutStr & "<TR><TD>&nbsp;</TD></TR>"										
					iOutStr = iOutStr & "</TABLE>"				
								
				End If

				lgObjRs3.Close
				set lgObjRs3 = Nothing			

				Response.Write "<Script Language=vbscript>"            & vbCr
				Response.Write " divDetachedFileInf.innerHTML = """ & iOutStr & """" & vbCr
				Response.Write "</Script>"                             & vbCr
				%>
			</TD>
		</TR>		
</TABLE>

<%
	End If
%>


<%
	End If
	
	Call SubCloseCommandObject(lgObjComm)    
	Call SubCloseDB(lgObjConn)
	
End If 

Function GetSpecialFolder(iDx)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   GetSpecialFolder = pfile.GetSpecialFolder(CInt(iDx))   
   Set pfile = Nothing
End Function
		
Function CreateFolder(iTempPath)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   Call pfile.CreateFolder(iTempPath)   
   Set pfile = Nothing
End Function

%>

<IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
<IFRAME NAME="FR_ATTWIZ" SRC="FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 FRAMEBORDER=0></IFRAME><BR>
</BODY>


</HTML>

