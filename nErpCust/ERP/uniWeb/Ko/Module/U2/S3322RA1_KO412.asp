<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<HTML>
<HEAD>
<STYLE>

TEXTAREA
{
    BORDER-RIGHT: 1px solid;
    BORDER-TOP: 1px solid;
    BORDER-LEFT: 1px solid;
    WIDTH: 100%;
    BORDER-BOTTOM: 1px solid;
    HEIGHT: 100%;
    BACKGROUND-COLOR: lightgrey
 

}
</STYLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<SCRIPT LANGUAGE="javascript" SRC="./js/btn.js"></SCRIPT>
<Script Language="VBScript">

CONST BIZ_PGM_ID = "S3322MB1_KO412.ASP"
Const BIZ_PGM_ID1 = "frviewBiz.asp" 

Sub window_onload()
	
End Sub

Sub window_onblur()
	//Call lostfocus
End Sub

Sub document_onclick()
	//Call getfocus
End Sub

Sub getfocus()
	//myView.bgColor = "navy"
	For i = 0 to myFnt.length-1
		myFnt(i).color = "white"
	Next
	Call Parent.frTitle.lostfocus
End Sub

Sub lostfocus()
	//myView.bgColor = "D1E8F9"
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

<!-- #Include file="../../inc/UNI2KCMCom.inc" -->	

<%


'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    

Dim lgObjConn, lgObjComm,lgObjRs, lgObjRs2, lgObjRs3, lgStrSQL
Dim strProject_Code, strReportNo, blnStop
Dim strins_person, strTitle, strReport_text, strinsrt_user_id, strinsrt_dt

Dim strData
Dim arrFileInfo
Dim iDx
Dim iCount
Dim tmpString
Dim tmpString1
Dim tmpString2
Dim iTempPath
dim filePath

Dim fso
Dim strSystemFolder

strProject_Code = FilterVar(Request("Project_code"), "''", "S") 
strReportNo = FilterVar(Request("Report_no"), "''", "S") 

on error Resume Next
strins_person = "" : strTitle = "": strReport_text = "": strinsrt_user_id = "" : strinsrt_dt = ""
blnStop = False

	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)	

If strProject_Code <> "" Then
  
	filePath=server.MapPath (".") & "\files\"
    	
	Redim arrFileInfo(100)
	
	lgStrSQL = "SELECT ins_user, report_no, insrt_user_id, REPORT_TEXT, insrt_dt FROM S_PRJ_REPORT_HDR_KO412 "
	lgStrSQL =lgStrSQL & " WHERE project_code=" & strProject_Code
	lgStrSQL =lgStrSQL & " AND report_no=" & strReportNo
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	

		'Call DisplayMsgBox("210037", vbInformation, "", "", I_MKSCRIPT)      '☜ : 게시물에 해당하는 자료가 존재하지 않습니다.

		'Response.Write "<Script Language=vbscript>"			& vbCr
		'Response.write "	Parent.window.frames(""frTitle"").document.URL = ""S3322MA1_KO412_frtitle.asp?page=""" & "& CStr(parent.window.frames(""frTitle"").intNowPage)" & vbCr
		'Response.Write "</Script>" & vbCr
		'Response.End					
	else

		strins_person = "" & lgObjRs("ins_user")
		strTitle = "" & lgObjRs("report_no")
		strinsrt_dt = "" & lgObjRs("insrt_dt")
		strReport_text = "" & lgObjRs("REPORT_TEXT")
		strinsrt_user_id = "" & lgObjRs("insrt_user_id")
		
	end if
	
end if	
		

%>

</HEAD>

<BODY bgcolor="#F4F3F3" topmargin=0>
<FORM NAME=frm1  METHOD="POST">

<TABLE  CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0 border=1 >
	<TR HEIGHT=100%>
		<TD WIDTH=100% CLASS="Tab11">

		<TABLE  CLASS="BasicTB" valign=top CELLSPACING=0>
				<TR>
					<TD  HEIGHT=5 WIDTH=100%></TD>
				</TR>
				
					   <!--TR>
							<TD CLASS="TD5" NOWRAP>작성자</TD>
							<TD CLASS="TD6" NOWRAP ><%= strins_person %><input type=hidden name=ins_Person value="" ></td>
						
							
							<TD CLASS="TD5" NOWRAP>제목</TD>
							<TD CLASS="TD6" NOWRAP ><%= strtitle %><input type=hidden name=title value="" ></td>
							
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>아이디</TD>
							<TD CLASS="TD6" NOWRAP><%= strINsrt_user_id %><input type=hidden name=insrt_user_Id value="" ></td>
							<TD CLASS="TD5" NOWRAP>작성일</TD>
							<TD CLASS="TD6" NOWRAP><%= strinsrt_Dt %><input type=hidden name=insrt_dt value="" ></TD>
						</TR-->
	
					<%

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
			
				'lgStrSQL = "SELECT REPORT_NM,REPORT_ID,REPORT_SEQ_NO, REPORT_PATH FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE = " & strProject_Code & " ORDER BY INSRT_DT ASC"
				lgStrSQL = "SELECT REPORT_NM, REPORT_ID,REPORT_SEQ_NO,REPORT_PATH FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE = " & strProject_Code & " AND REPORT_NO = " & strReportNo & " ORDER BY PROJECT_CODE "

				If 	FncOpenRs("R",lgObjConn,lgObjRs3,lgStrSQL,"X","X") = False Then  
				
				else
				%>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD" >
					<TABLE  CLASS="BasicTB"  CELLSPACING=0>
					
					<TR>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></td>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>첨부파일</TD>
						<TD CLASS="TD6" NOWRAP >
						
						<!--실행모드 
						
						<SELECT NAME=cboFileMode>
						<OPTION value="W">저장</OPTION>
						<OPTION value="F" selected>보기</OPTION></SELECT-->

						<%
									If lgObjRs3.BOF or lgObjRs3.EOF Then
										Response.Write  "첨부된 파일이 없습니다."
									Else
										'첨부된 파일 정보를 테이블로 조립하는 부분
										Dim iOutStr
										iDx = 0

									    Do Until lgObjRs3.EOF
									    
											strData = ""
											strData = "1" & lgObjRs3(0) & "" & lgObjRs3(3) & lgObjRs3(0) & "" & iTempPath & lgObjRs3(1) & "101344919970601092656N0NNFFY00YI"
									    
											'iOutStr = iOutStr & "&nbsp;<A onclick='VBSCRIPT:GETFILE(" & iDx & ")'>" & lgObjRs3(0) & "</A>&nbsp;&nbsp;&nbsp;&nbsp;"
											'''iOutStr = iOutStr & "&nbsp;<A onclick='VBSCRIPT:GETFILE(""" & strData & """)'>" & lgObjRs3(0) & "</A><br>"
											'iOutStr = iOutStr & "&nbsp;<A href='"&"./files/"&lgObjRs3(1)&"' target = '_parent'>" & lgObjRs3(0) & "</A><br>"
											iOutStr = iOutStr & "&nbsp;<A href='"&"./files/"&lgObjRs3(1)&"'>" & lgObjRs3(0) & "</A><br>"
											lgObjRs3.MoveNext
											iDx = iDx + 1
										Loop

									End If

									lgObjRs3.Close
									set lgObjRs3 = Nothing
										Response.Write iOutStr
									%>
									
						</td>
						<TD CLASS="TD6" width=1 NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
							
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<%end if%>

				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%"  CLASS="BasicTB" CELLSPACING=0>
							<TR>
								<TD HEIGHT="100%">
									<textarea tabIndex="-1"  name="REPORT_TEXT"  readonly><%= strReport_text %></textarea>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
	</td></tr></table>		
	
<%
'===============================================================
' file 아이디 vbs로 구현하기
'===============================================================

	lgStrSQL = "SELECT REPORT_NM, REPORT_ID, REPORT_PATH FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE = " & strProject_Code & " ORDER BY PROJECT_CODE "

	If 	FncOpenRs("R",lgObjConn,lgObjRs2,lgStrSQL,"X","X") <> False Then                    'If data not exists	
%>
	<script language="vbscript">
	Function GetFile(iDx)
	
		Dim strVal
		
		On Error Resume Next
		
	  <%
	  
'		iDx = 0
'		Do While Not lgObjRs2.EOF
'			strData = "1" & lgObjRs2(0) & "" & lgObjRs2(2) & lgObjRs2(0) & "" & iTempPath & lgObjRs2(1) & "101344919970601092656N0NNFFY00YI"
'			arrFileInfo(iDx) = strData
'			
'			lgObjRs2.MoveNext
'			iDx = iDx + 1
'			response.write " strVal = BIZ_PGM_ID1&""?FileInfo=" & arrFileInfo(0) & """" & vbCr
'		Loop

		%>

		strVal = BIZ_PGM_ID1 & "?FileInfo=" & iDx
		strVal = strVal & "&FileMode=" & "F"	'Trim(cboFileMode.value)

		Call RunMyBizASP(MyBizASP, strVal)

	End Function
	</script>

<% End if

'===============================================================	
	Call SubCloseCommandObject(lgObjComm)    
	Call SubCloseDB(lgObjConn)

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

<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=800 FRAMEBORDER=0 SCROLLING=NO framespacing=0 tabindex=-1></IFRAME>
<IFRAME NAME="FR_ATTWIZ" SRC="FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 FRAMEBORDER=0></IFRAME><BR>

</BODY>

</HTML>

